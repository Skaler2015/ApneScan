<?php
/**
 * ApneScan Phone Relay — "Scan from phone, ANYWHERE" (internet upload bridge)
 * ---------------------------------------------------------------------------
 * Kaise kaam karta hai:
 *   PC app  --api=create-->  yahan session banta hai (token + secret key)
 *   Phone   --QR se ?u=TOKEN--> upload page --api=upload--> file yahan aati hai
 *   PC app  --api=status/take (key ke saath)--> file utha kar TURANT delete
 *
 * Suraksha:
 *   - TOKEN sirf UPLOAD karne deta hai; files WAPAS sirf KEY (jo QR me nahi
 *     hota, sirf PC app ke paas) se milti hain.
 *   - Har session apne aap expire (5-60 min) + files auto-delete.
 *   - MIME + extension whitelist, size/count limits, per-IP rate limit.
 *   - uploads folder web se poori tarah band (.htaccess deny + .bin naam).
 *   - HTTPS host (Hostinger) — transit me encrypted.
 *
 * Install: is file ko stats.php ke saath (status.apnesoft.com) upload karo.
 * App me URL: https://status.apnesoft.com/phone_relay.php
 */

error_reporting(0);
header('X-Content-Type-Options: nosniff');

$DIR       = __DIR__ . '/phone_uploads';
$MAX_FILE  = 25 * 1024 * 1024;    // ek file
$MAX_TOTAL = 120 * 1024 * 1024;   // poore session ka jod
$MAX_FILES = 40;                  // session me files
$TTL_MIN   = 60; $TTL_MAX = 3600; // 1 min – 60 min
$RATE_UP_PER_HOUR = 200;          // per-IP uploads/ghanta
$EXT_OK = array('jpg','jpeg','png','webp','heic','heif','pdf','tif','tiff');
$MIME_OK = array('image/jpeg','image/png','image/webp','image/heic','image/heif',
                 'image/heic-sequence','application/pdf','image/tiff','application/octet-stream');

// ---------- storage init ----------
if (!is_dir($DIR)) { @mkdir($DIR, 0755, true); }
if (!is_file($DIR.'/.htaccess')) {
    @file_put_contents($DIR.'/.htaccess',
        "Require all denied\nDeny from all\nphp_flag engine off\n");
}
if (!is_file($DIR.'/index.html')) { @file_put_contents($DIR.'/index.html', ''); }

function jout($a, $code = 200) {
    http_response_code($code);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode($a); exit;
}
function clean_token($t) { return substr(preg_replace('/[^A-Z0-9]/', '', strtoupper((string)$t)), 0, 32); }
function clean_key($k)   { return substr(preg_replace('/[^a-f0-9]/', '', strtolower((string)$k)), 0, 64); }
function sess_file($t)   { global $DIR; return $DIR . '/s_' . $t . '.json'; }

function sess_load($t) {
    $f = sess_file($t); if (!is_file($f)) return null;
    $s = json_decode((string)@file_get_contents($f), true);
    if (!is_array($s)) return null;
    if (time() - intval($s['created']) > intval($s['ttl'])) return null;   // expired
    if (!empty($s['stopped'])) return null;
    return $s;
}
function sess_save($t, $s) {
    $f = sess_file($t); $tmp = $f . '.tmp';
    @file_put_contents($tmp, json_encode($s), LOCK_EX);
    @rename($tmp, $f);
}
function sess_destroy($t) {
    global $DIR; $s = null;
    $f = sess_file($t);
    $j = json_decode((string)@file_get_contents($f), true);
    if (is_array($j) && !empty($j['files'])) {
        foreach ($j['files'] as $fm) { @unlink($DIR . '/' . basename($fm['p'])); }
    }
    if (is_array($j) && !empty($j['pcfiles'])) {   // (v277) PC->phone files bhi hatao
        foreach ($j['pcfiles'] as $fm) { @unlink($DIR . '/' . basename($fm['p'])); }
    }
    @unlink($f);
}

// ---------- garbage-collect (har request par, sasta) ----------
function gc() {
    global $DIR;
    $now = time();
    foreach ((array)@glob($DIR . '/s_*.json') as $f) {
        $s = json_decode((string)@file_get_contents($f), true);
        $dead = !is_array($s) || ($now - intval($s['created']) > intval($s['ttl']) + 120)
                || !empty($s['stopped']);
        if ($dead) {
            if (is_array($s) && !empty($s['files'])) {
                foreach ($s['files'] as $fm) { @unlink($DIR . '/' . basename($fm['p'])); }
            }
            if (is_array($s) && !empty($s['pcfiles'])) {
                foreach ($s['pcfiles'] as $fm) { @unlink($DIR . '/' . basename($fm['p'])); }
            }
            @unlink($f);
        }
    }
    // 3 ghante se purani koi bhi upload-file (safety net) — dono taraf (f_/g_)
    foreach (array_merge((array)@glob($DIR . '/f_*.bin'), (array)@glob($DIR . '/g_*.bin')) as $f) {
        if ($now - (int)@filemtime($f) > 3 * 3600) @unlink($f);
    }
}
gc();

function rate_ok($kind, $limit) {
    global $DIR;
    $ip = isset($_SERVER['REMOTE_ADDR']) ? $_SERVER['REMOTE_ADDR'] : '?';
    $f = $DIR . '/rate_' . $kind . '_' . md5($ip) . '.json';
    $r = json_decode((string)@file_get_contents($f), true);
    $hr = date('YmdH');
    if (!is_array($r) || $r['h'] !== $hr) $r = array('h' => $hr, 'n' => 0);
    $r['n']++;
    @file_put_contents($f, json_encode($r));
    return $r['n'] <= $limit;
}

$api = isset($_REQUEST['api']) ? $_REQUEST['api'] : '';

// ================= API: create (PC app) =================
if ($api === 'create') {
    if (!rate_ok('create', 30)) jout(array('ok'=>0,'err'=>'busy'), 429);
    $t = clean_token(isset($_POST['t']) ? $_POST['t'] : '');
    $k = clean_key(isset($_POST['k']) ? $_POST['k'] : '');
    $ttl = max($GLOBALS['TTL_MIN'], min($GLOBALS['TTL_MAX'], intval(isset($_POST['ttl']) ? $_POST['ttl'] : 1800)));
    if (strlen($t) < 12 || strlen($k) < 32) jout(array('ok'=>0,'err'=>'bad token'), 400);
    if (is_file(sess_file($t))) jout(array('ok'=>0,'err'=>'exists'), 409);
    sess_save($t, array('created'=>time(), 'ttl'=>$ttl, 'kh'=>hash('sha256', $k),
                        'files'=>array(), 'pcfiles'=>array(),
                        'taken'=>0, 'seen'=>0, 'stopped'=>0));
    jout(array('ok'=>1, 'ttl'=>$ttl));
}

// ================= API: upload (phone) =================
if ($api === 'upload') {
    $t = clean_token(isset($_REQUEST['t']) ? $_REQUEST['t'] : '');
    $s = sess_load($t);
    if ($s === null) jout(array('ok'=>0,'err'=>'expired'), 410);
    if (!rate_ok('up', $RATE_UP_PER_HOUR)) jout(array('ok'=>0,'err'=>'busy'), 429);
    if (count($s['files']) >= $MAX_FILES) jout(array('ok'=>0,'err'=>'too many files'), 413);
    if (!isset($_FILES['file']) || !is_uploaded_file($_FILES['file']['tmp_name']))
        jout(array('ok'=>0,'err'=>'no file'), 400);
    $sz = intval($_FILES['file']['size']);
    if ($sz <= 0 || $sz > $MAX_FILE) jout(array('ok'=>0,'err'=>'file too large'), 413);
    $tot = 0; foreach ($s['files'] as $fm) $tot += intval($fm['s']);
    if ($tot + $sz > $MAX_TOTAL) jout(array('ok'=>0,'err'=>'session full'), 413);
    $name = preg_replace('/[^A-Za-z0-9._ -]/', '_', (string)$_FILES['file']['name']);
    $name = substr($name !== '' ? $name : 'file', 0, 80);
    $ext = strtolower(pathinfo($name, PATHINFO_EXTENSION));
    if (!in_array($ext, $EXT_OK)) jout(array('ok'=>0,'err'=>'unsupported type'), 415);
    $mime = '';
    if (function_exists('finfo_open')) {
        $fi = finfo_open(FILEINFO_MIME_TYPE);
        if ($fi) { $mime = (string)finfo_file($fi, $_FILES['file']['tmp_name']); finfo_close($fi); }
    }
    if ($mime !== '' && !in_array($mime, $MIME_OK)) jout(array('ok'=>0,'err'=>'unsupported type'), 415);
    $id = bin2hex(function_exists('random_bytes') ? random_bytes(8) : md5(uniqid('', true)));
    $p = 'f_' . $t . '_' . $id . '.bin';
    if (!@move_uploaded_file($_FILES['file']['tmp_name'], $DIR . '/' . $p))
        jout(array('ok'=>0,'err'=>'server error'), 500);
    $s['files'][] = array('id'=>$id, 'n'=>$name, 's'=>$sz, 'p'=>$p, 't'=>time());
    sess_save($t, $s);
    jout(array('ok'=>1, 'id'=>$id, 'count'=>count($s['files'])));
}

// ================= API: pstat (phone — halka status) =================
if ($api === 'pstat') {
    $t = clean_token(isset($_REQUEST['t']) ? $_REQUEST['t'] : '');
    $s = sess_load($t);
    if ($s === null) jout(array('ok'=>0,'expired'=>1));
    jout(array('ok'=>1, 'got'=>intval($s['taken']),
               'pc'=>(time() - intval($s['seen']) < 12 ? 1 : 0),
               'left'=>max(0, intval($s['created']) + intval($s['ttl']) - time())));
}

// ================= API: status (PC app, key zaroori) =================
if ($api === 'status') {
    $t = clean_token(isset($_REQUEST['t']) ? $_REQUEST['t'] : '');
    $k = clean_key(isset($_REQUEST['k']) ? $_REQUEST['k'] : '');
    $s = sess_load($t);
    if ($s === null) jout(array('ok'=>0,'err'=>'expired'), 410);
    if (!hash_equals($s['kh'], hash('sha256', $k))) jout(array('ok'=>0,'err'=>'denied'), 403);
    $s['seen'] = time(); sess_save($t, $s);
    $out = array();
    foreach ($s['files'] as $fm) $out[] = array('id'=>$fm['id'], 'name'=>$fm['n'], 'size'=>intval($fm['s']));
    jout(array('ok'=>1, 'files'=>$out, 'taken'=>intval($s['taken']),
               'left'=>max(0, intval($s['created']) + intval($s['ttl']) - time())));
}

// ================= API: take (PC app — file do, phir DELETE) =================
if ($api === 'take') {
    $t = clean_token(isset($_REQUEST['t']) ? $_REQUEST['t'] : '');
    $k = clean_key(isset($_REQUEST['k']) ? $_REQUEST['k'] : '');
    $id = substr(preg_replace('/[^a-f0-9]/', '', (string)(isset($_REQUEST['id']) ? $_REQUEST['id'] : '')), 0, 32);
    $s = sess_load($t);
    if ($s === null) jout(array('ok'=>0,'err'=>'expired'), 410);
    if (!hash_equals($s['kh'], hash('sha256', $k))) jout(array('ok'=>0,'err'=>'denied'), 403);
    $found = null; $rest = array();
    foreach ($s['files'] as $fm) { if ($fm['id'] === $id) $found = $fm; else $rest[] = $fm; }
    if ($found === null) jout(array('ok'=>0,'err'=>'not found'), 404);
    $fp = $DIR . '/' . basename($found['p']);
    if (!is_file($fp)) jout(array('ok'=>0,'err'=>'not found'), 404);
    $s['files'] = $rest; $s['taken'] = intval($s['taken']) + 1; $s['seen'] = time();
    sess_save($t, $s);
    header('Content-Type: application/octet-stream');
    header('Content-Length: ' . filesize($fp));
    header('X-File-Name: ' . $found['n']);
    readfile($fp);
    @unlink($fp);
    exit;
}

// ================= API: pcsend (PC app -> phone, KEY zaroori) =================
// PC bharosemand hai (KEY uske paas), isliye KOI BHI file-type bhej sakta hai.
if ($api === 'pcsend') {
    $t = clean_token(isset($_REQUEST['t']) ? $_REQUEST['t'] : '');
    $k = clean_key(isset($_REQUEST['k']) ? $_REQUEST['k'] : '');
    $s = sess_load($t);
    if ($s === null) jout(array('ok'=>0,'err'=>'expired'), 410);
    if (!hash_equals($s['kh'], hash('sha256', $k))) jout(array('ok'=>0,'err'=>'denied'), 403);
    if (!isset($s['pcfiles']) || !is_array($s['pcfiles'])) $s['pcfiles'] = array();
    if (count($s['pcfiles']) >= $MAX_FILES) jout(array('ok'=>0,'err'=>'too many files'), 413);
    if (!isset($_FILES['file']) || !is_uploaded_file($_FILES['file']['tmp_name']))
        jout(array('ok'=>0,'err'=>'no file'), 400);
    $sz = intval($_FILES['file']['size']);
    if ($sz <= 0 || $sz > $MAX_FILE) jout(array('ok'=>0,'err'=>'file too large'), 413);
    $tot = 0; foreach ($s['pcfiles'] as $fm) $tot += intval($fm['s']);
    if ($tot + $sz > $MAX_TOTAL) jout(array('ok'=>0,'err'=>'session full'), 413);
    $name = preg_replace('/[^A-Za-z0-9._ ()-]/', '_', (string)$_FILES['file']['name']);
    $name = substr($name !== '' ? $name : 'file', 0, 90);
    $id = bin2hex(function_exists('random_bytes') ? random_bytes(8) : md5(uniqid('', true)));
    $p = 'g_' . $t . '_' . $id . '.bin';   // g_ = PC->phone (f_ = phone->PC)
    if (!@move_uploaded_file($_FILES['file']['tmp_name'], $DIR . '/' . $p))
        jout(array('ok'=>0,'err'=>'server error'), 500);
    $s['pcfiles'][] = array('id'=>$id, 'n'=>$name, 's'=>$sz, 'p'=>$p, 't'=>time());
    $s['seen'] = time();
    sess_save($t, $s);
    jout(array('ok'=>1, 'id'=>$id, 'count'=>count($s['pcfiles'])));
}

// ================= API: pclist (phone — PC ki bheji files ki soochi) =========
if ($api === 'pclist') {
    $t = clean_token(isset($_REQUEST['t']) ? $_REQUEST['t'] : '');
    $s = sess_load($t);
    if ($s === null) jout(array('ok'=>0,'expired'=>1));
    $out = array();
    if (!empty($s['pcfiles'])) {
        foreach ($s['pcfiles'] as $fm)
            $out[] = array('id'=>$fm['id'], 'name'=>$fm['n'], 'size'=>intval($fm['s']));
    }
    jout(array('ok'=>1, 'files'=>$out));
}

// ================= API: pcget (phone — PC ki bheji file download) ============
if ($api === 'pcget') {
    $t = clean_token(isset($_REQUEST['t']) ? $_REQUEST['t'] : '');
    $id = substr(preg_replace('/[^a-f0-9]/', '', (string)(isset($_REQUEST['id']) ? $_REQUEST['id'] : '')), 0, 32);
    $s = sess_load($t);
    if ($s === null) jout(array('ok'=>0,'err'=>'expired'), 410);
    $found = null;
    if (!empty($s['pcfiles'])) {
        foreach ($s['pcfiles'] as $fm) { if ($fm['id'] === $id) { $found = $fm; break; } }
    }
    if ($found === null) jout(array('ok'=>0,'err'=>'not found'), 404);
    $fp = $DIR . '/' . basename($found['p']);
    if (!is_file($fp)) jout(array('ok'=>0,'err'=>'not found'), 404);
    // phone ko download-attachment de do (server par rehne do — dubara chahiye
    // to phir mile; session expire par apne aap delete).
    $dl = preg_replace('/[^A-Za-z0-9._ ()-]/', '_', (string)$found['n']);
    header('Content-Type: application/octet-stream');
    header('Content-Length: ' . filesize($fp));
    header('Content-Disposition: attachment; filename="' . $dl . '"');
    readfile($fp);
    exit;
}

// ================= API: stop (PC app) =================
if ($api === 'stop') {
    $t = clean_token(isset($_REQUEST['t']) ? $_REQUEST['t'] : '');
    $k = clean_key(isset($_REQUEST['k']) ? $_REQUEST['k'] : '');
    $f = sess_file($t);
    $j = json_decode((string)@file_get_contents($f), true);
    if (is_array($j) && isset($j['kh']) && hash_equals($j['kh'], hash('sha256', $k))) {
        sess_destroy($t);
        jout(array('ok'=>1));
    }
    jout(array('ok'=>0,'err'=>'denied'), 403);
}

// ================= PHONE UPLOAD PAGE (?u=TOKEN) =================
$u = clean_token(isset($_GET['u']) ? $_GET['u'] : '');
header('Content-Type: text/html; charset=utf-8');
$valid = ($u !== '' && sess_load($u) !== null);
?><!doctype html>
<html lang="hi"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>Send Files to ApneScan</title>
<meta name="robots" content="noindex">
<style>
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
body{margin:0;font-family:system-ui,'Segoe UI',Roboto,Arial;background:#0b1220;color:#e6edf7;
  min-height:100vh;display:flex;flex-direction:column;align-items:center;padding:18px 14px 40px}
.card{width:100%;max-width:430px;background:#111a2e;border:1px solid #22304d;border-radius:20px;
  padding:20px 18px;box-shadow:0 18px 50px rgba(0,0,0,.45)}
h1{font-size:19px;margin:2px 0 4px;display:flex;align-items:center;gap:8px}
.sub{color:#93a4c3;font-size:12.5px;margin-bottom:14px}
.btn{display:flex;align-items:center;gap:12px;width:100%;padding:15px 16px;margin:8px 0;border:none;
  border-radius:14px;background:#1b2946;color:#fff;font-size:15.5px;font-weight:700;cursor:pointer;text-align:left}
.btn:active{transform:scale(.985)}
.btn .ic{font-size:22px}
.btn.cam{background:linear-gradient(135deg,#2a78d6,#1f5fb0)}
.opt{display:flex;align-items:center;gap:8px;font-size:12.5px;color:#93a4c3;margin:10px 2px}
.bar{height:8px;background:#1b2946;border-radius:6px;overflow:hidden;margin-top:6px}
.bar i{display:block;height:100%;width:0%;background:linear-gradient(90deg,#2a78d6,#22b8a8);transition:width .2s}
.item{background:#0d1526;border:1px solid #22304d;border-radius:12px;padding:10px 12px;margin:7px 0;font-size:12.5px}
.item .nm{display:flex;justify-content:space-between;gap:8px}
.ok{color:#3fd68f;font-weight:700}.bad{color:#ff7d7d;font-weight:700}
.retry{background:none;border:1px solid #3b4c74;color:#9db4e0;border-radius:8px;padding:3px 10px;font-size:11.5px;cursor:pointer}
.pc{font-size:12px;color:#3fd68f;font-weight:700;margin-top:10px;min-height:16px}
.exp{font-size:11.5px;color:#93a4c3;margin-top:3px}
.dead{background:#2a1520;border:1px solid #5d2b3a;border-radius:14px;padding:16px;color:#ffb3c0;font-size:14px;line-height:1.6}
.done{display:none;text-align:center;padding:14px 0 4px}
.done .big{font-size:40px}
.foot{color:#5b6b8c;font-size:10.5px;margin-top:16px;text-align:center}
</style></head><body>
<div class="card">
<h1>📄 Send Files to ApneScan</h1>
<div class="sub">File turant aapke computer par pahunchti hai aur server se DELETE ho jaati hai.</div>
<?php if (!$valid) { ?>
  <div class="dead"><b>⛔ Link expire ho gayi ya galat hai.</b><br>
  Computer par ApneScan me "📱 Phone se scan" dobara khol kar naya QR scan karo.</div>
<?php } else { ?>
  <button class="btn cam" onclick="pick('cam')"><span class="ic">📷</span> Capture Photo</button>
  <button class="btn" onclick="pick('gal')"><span class="ic">🖼</span> Choose from Gallery</button>
  <button class="btn" onclick="pick('pdf')"><span class="ic">📄</span> Upload PDF</button>
  <button class="btn" onclick="pick('any')"><span class="ic">📁</span> Browse Files</button>
  <label class="opt"><input type="checkbox" id="optim" checked>
    Optimized upload (photo chhoti karke tez bhejo — quality print-layak)</label>
  <input type="file" id="fcam" accept="image/*" capture="environment" style="display:none">
  <input type="file" id="fgal" accept="image/*" multiple style="display:none">
  <input type="file" id="fpdf" accept="application/pdf" multiple style="display:none">
  <input type="file" id="fany" accept=".jpg,.jpeg,.png,.webp,.heic,.pdf,.tif,.tiff" multiple style="display:none">
  <div id="list"></div>
  <!-- (v277) Computer se aayi files — phone par download -->
  <div id="frompc" style="display:none;margin-top:14px">
    <div style="font-size:13.5px;font-weight:700;color:#9db4e0;margin:6px 2px">⬇️ Computer se aayi files</div>
    <div id="pclist"></div>
  </div>
  <div class="pc" id="pc"></div>
  <div class="exp" id="exp"></div>
  <div class="done" id="done"><div class="big">✅</div><b>Files Sent Successfully</b><br>
    <span style="color:#93a4c3;font-size:12.5px">Computer par pahunch gayi hain.</span><br><br>
    <button class="btn" style="justify-content:center" onclick="more()">➕ Send More</button></div>
<?php } ?>
<div class="foot">ApneScan · Secure temporary link · files auto-delete</div>
</div>
<script>
var T = <?php echo json_encode($u); ?>;
var sent = 0, active = 0;
function pick(k){ document.getElementById('f'+({cam:'cam',gal:'gal',pdf:'pdf',any:'any'})[k]).click(); }
['fcam','fgal','fpdf','fany'].forEach(function(id){
  var el = document.getElementById(id); if(!el) return;
  el.addEventListener('change', function(e){
    var fs = [].slice.call(e.target.files); e.target.value='';
    fs.forEach(function(f){ prep(f); });
  });
});
function prep(f){
  var isImg = /^image\//.test(f.type) && !/heic|heif/i.test(f.type);
  if (document.getElementById('optim').checked && isImg && f.size > 700*1024) {
    shrink(f, function(blob){ up(blob || f, f.name); });
  } else up(f, f.name);
}
function shrink(f, cb){
  try{
    var img = new Image(), url = URL.createObjectURL(f);
    img.onload = function(){
      var mx = 2400, w = img.width, h = img.height;
      if (Math.max(w,h) > mx){ var s = mx/Math.max(w,h); w = Math.round(w*s); h = Math.round(h*s); }
      var c = document.createElement('canvas'); c.width = w; c.height = h;
      c.getContext('2d').drawImage(img, 0, 0, w, h);
      URL.revokeObjectURL(url);
      c.toBlob(function(b){ cb(b); }, 'image/jpeg', 0.9);
    };
    img.onerror = function(){ URL.revokeObjectURL(url); cb(null); };
    img.src = url;
  }catch(e){ cb(null); }
}
function up(file, name, holder){
  var d = holder;
  if (!d){
    d = document.createElement('div'); d.className='item';
    d.innerHTML = '<div class="nm"><span>'+esc(name)+'</span><span class="st">⏳</span></div><div class="bar"><i></i></div>';
    var L = document.getElementById('list'); L.insertBefore(d, L.firstChild);
  }
  var st = d.querySelector('.st'), bar = d.querySelector('.bar i');
  st.textContent = '⏳'; active++;
  var fd = new FormData();
  fd.append('file', file, name || 'photo.jpg');
  var x = new XMLHttpRequest();
  x.open('POST', '?api=upload&t=' + T);
  x.upload.onprogress = function(ev){ if (ev.total) bar.style.width = Math.round(ev.loaded*100/ev.total) + '%'; };
  x.onload = function(){
    active--;
    if (x.status === 200){ st.textContent=''; st.className='st ok'; st.textContent='✅ Sent'; sent++; fin(); }
    else {
      var msg = {410:'Link expire ho gayi',413:'File badi hai',415:'Ye type support nahi',429:'Server busy — thoda ruk kar'}[x.status] || 'Upload fail';
      fail(d, st, msg, file, name);
    }
  };
  x.onerror = function(){ active--; fail(d, st, 'Network error', file, name); };
  x.send(fd);
}
function fail(d, st, msg, file, name){
  st.className='st bad'; st.innerHTML = '❌ ' + msg + ' ';
  var b = document.createElement('button'); b.className='retry'; b.textContent='↻ Retry';
  b.onclick = function(){ b.remove(); up(file, name, d); };
  st.appendChild(b);
}
function fin(){ if (active === 0 && sent > 0) document.getElementById('done').style.display='block'; }
function more(){ document.getElementById('done').style.display='none'; }
function esc(s){ return String(s).replace(/[&<>]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c]; }); }
function mmss(s){ return Math.floor(s/60)+':'+('0'+(s%60)).slice(-2); }
function poll(){
  if (!T) return;
  var x = new XMLHttpRequest();
  x.open('GET', '?api=pstat&t=' + T + '&_=' + Date.now());
  x.onload = function(){
    try{
      var r = JSON.parse(x.responseText);
      if (!r.ok){ document.getElementById('pc').textContent=''; document.getElementById('exp').textContent='⛔ Link expire ho gayi'; return; }
      document.getElementById('pc').textContent = (r.pc ? '💻 Computer connected' : '') + (r.got ? ' · ' + r.got + ' file computer ne le li ✓' : '');
      document.getElementById('exp').textContent = '⏱ Link ' + mmss(r.left) + ' me expire hogi';
    }catch(e){}
    setTimeout(poll, 4000);
  };
  x.onerror = function(){ setTimeout(poll, 8000); };
  x.send();
}
poll();
// (v277) Computer se bheji files ki soochi — download buttons
function fmtSize(n){ return n<1048576 ? Math.round(n/1024)+' KB' : (n/1048576).toFixed(1)+' MB'; }
function pcpoll(){
  if (!T) return;
  var x = new XMLHttpRequest();
  x.open('GET', '?api=pclist&t=' + T + '&_=' + Date.now());
  x.onload = function(){
    try{
      var r = JSON.parse(x.responseText);
      var box = document.getElementById('frompc'), L = document.getElementById('pclist');
      if (r.ok && r.files && r.files.length){
        box.style.display = 'block';
        var h = '';
        r.files.forEach(function(f){
          h += '<div class="item"><div class="nm"><span>'+esc(f.name)+
               '</span><a class="retry" style="text-decoration:none" href="?api=pcget&t='+T+
               '&id='+encodeURIComponent(f.id)+'" download="'+esc(f.name)+'">⬇ Download</a></div>'+
               '<div style="color:#93a4c3;font-size:11px;margin-top:3px">'+fmtSize(f.size)+'</div></div>';
        });
        L.innerHTML = h;
      } else { box.style.display = 'none'; }
    }catch(e){}
    setTimeout(pcpoll, 5000);
  };
  x.onerror = function(){ setTimeout(pcpoll, 9000); };
  x.send();
}
pcpoll();
</script></body></html>
