<?php
/**
 * ApneScan Web Share — "Phone to Phone" (bina app ke, kisi bhi device se)
 * ---------------------------------------------------------------------------
 * ROOM model (symmetric): ek device room banata hai -> ek CODE + link milta
 * hai. Doosra device wahi link kholta / code daalta hai. DONO device us room
 * me koi bhi file daal/utar sakte hain (phone<->phone, phone<->PC). Sab file
 * ROOM expire hote hi delete. CODE hi shared secret hai (link me).
 *
 * Suraksha:
 *   - Room 30-60 min me apne aap expire + saari files delete.
 *   - Files '.bin' naam se, folder web se poori tarah band (.htaccess deny +
 *     php engine off) — kabhi execute nahi hotीं.
 *   - Size/count/rate limit. HTTPS host (transit encrypted).
 *
 * Install: is file ko stats.php/phone_relay.php ke saath (status.apnesoft.com)
 * upload karo. Website par link: https://status.apnesoft.com/share.php
 */

error_reporting(0);
header('X-Content-Type-Options: nosniff');
// (v285) apnescan.apnesoft.com/share (GitHub Pages) is tool ko IFRAME me dikhata
// hai — isliye apne domain se framing allow karo (X-Frame-Options hata kar CSP).
@header_remove('X-Frame-Options');
header("Content-Security-Policy: frame-ancestors 'self' https://apnescan.apnesoft.com https://apnesoft.com https://*.apnesoft.com");

$DIR       = __DIR__ . '/share_uploads';
$MAX_FILE  = 200 * 1024 * 1024;   // ek file
$MAX_TOTAL = 800 * 1024 * 1024;   // poore room ka jod
$MAX_FILES = 60;                  // room me files
$TTL       = 45 * 60;             // room ki umar (45 min)
$RATE_PER_HOUR = 400;             // per-IP uploads/ghanta
$CODE_CHARS = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';  // bina ambiguous (I,L,O,0,1)

if (!is_dir($DIR)) { @mkdir($DIR, 0755, true); }
if (!is_file($DIR.'/.htaccess')) {
    @file_put_contents($DIR.'/.htaccess', "Require all denied\nDeny from all\nphp_flag engine off\n");
}
if (!is_file($DIR.'/index.html')) { @file_put_contents($DIR.'/index.html', ''); }

function jout($a, $code = 200) {
    http_response_code($code);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode($a); exit;
}
function clean_code($c) { return substr(preg_replace('/[^A-Z0-9]/', '', strtoupper((string)$c)), 0, 12); }
function room_file($c) { global $DIR; return $DIR . '/r_' . $c . '.json'; }

function room_load($c) {
    $f = room_file($c); if (!is_file($f)) return null;
    $s = json_decode((string)@file_get_contents($f), true);
    if (!is_array($s)) return null;
    if (time() - intval($s['created']) > intval($s['ttl'])) return null;   // expired
    return $s;
}
function room_save($c, $s) {
    $f = room_file($c); $tmp = $f . '.tmp';
    @file_put_contents($tmp, json_encode($s), LOCK_EX);
    @rename($tmp, $f);
}

// ---------- garbage-collect (har request par, sasta) ----------
function gc_rooms() {
    global $DIR; $now = time();
    foreach ((array)@glob($DIR . '/r_*.json') as $f) {
        $s = json_decode((string)@file_get_contents($f), true);
        $dead = !is_array($s) || ($now - intval($s['created']) > intval($s['ttl']) + 120);
        if ($dead) {
            if (is_array($s) && !empty($s['files'])) {
                foreach ($s['files'] as $fm) { @unlink($DIR . '/' . basename($fm['p'])); }
            }
            @unlink($f);
        }
    }
    foreach ((array)@glob($DIR . '/sf_*.bin') as $f) {   // safety net (3 ghante)
        if ($now - (int)@filemtime($f) > 3 * 3600) @unlink($f);
    }
}
gc_rooms();

function rate_ok($limit) {
    global $DIR;
    $ip = isset($_SERVER['REMOTE_ADDR']) ? $_SERVER['REMOTE_ADDR'] : '?';
    $f = $DIR . '/rate_' . md5($ip) . '.json';
    $r = json_decode((string)@file_get_contents($f), true);
    $hr = date('YmdH');
    if (!is_array($r) || $r['h'] !== $hr) $r = array('h' => $hr, 'n' => 0);
    $r['n']++;
    @file_put_contents($f, json_encode($r));
    return $r['n'] <= $limit;
}

function gen_code() {
    global $CODE_CHARS;
    for ($tries = 0; $tries < 20; $tries++) {
        $c = '';
        for ($i = 0; $i < 6; $i++) {
            $c .= $CODE_CHARS[random_int(0, strlen($CODE_CHARS) - 1)];
        }
        if (!is_file(room_file($c))) return $c;
    }
    return null;
}

$api = isset($_REQUEST['api']) ? $_REQUEST['api'] : '';

// ================= API: create room =================
if ($api === 'create') {
    if (!rate_ok(60)) jout(array('ok'=>0,'err'=>'busy'), 429);
    $c = gen_code();
    if ($c === null) jout(array('ok'=>0,'err'=>'server busy'), 500);
    room_save($c, array('created'=>time(), 'ttl'=>$TTL, 'files'=>array()));
    jout(array('ok'=>1, 'code'=>$c, 'ttl'=>$TTL));
}

// ================= API: upload =================
if ($api === 'up') {
    $c = clean_code(isset($_REQUEST['r']) ? $_REQUEST['r'] : '');
    $s = room_load($c);
    if ($s === null) jout(array('ok'=>0,'err'=>'expired'), 410);
    if (!rate_ok($RATE_PER_HOUR)) jout(array('ok'=>0,'err'=>'busy'), 429);
    if (!isset($s['files']) || !is_array($s['files'])) $s['files'] = array();
    if (count($s['files']) >= $MAX_FILES) jout(array('ok'=>0,'err'=>'too many files'), 413);
    if (!isset($_FILES['file']) || !is_uploaded_file($_FILES['file']['tmp_name']))
        jout(array('ok'=>0,'err'=>'no file'), 400);
    $sz = intval($_FILES['file']['size']);
    if ($sz <= 0 || $sz > $MAX_FILE) jout(array('ok'=>0,'err'=>'file too large'), 413);
    $tot = 0; foreach ($s['files'] as $fm) $tot += intval($fm['s']);
    if ($tot + $sz > $MAX_TOTAL) jout(array('ok'=>0,'err'=>'room full'), 413);
    $name = preg_replace('/[^A-Za-z0-9._ ()-]/', '_', (string)$_FILES['file']['name']);
    $name = substr($name !== '' ? $name : 'file', 0, 90);
    $id = bin2hex(random_bytes(8));
    $p = 'sf_' . $c . '_' . $id . '.bin';
    if (!@move_uploaded_file($_FILES['file']['tmp_name'], $DIR . '/' . $p))
        jout(array('ok'=>0,'err'=>'server error'), 500);
    $s['files'][] = array('id'=>$id, 'n'=>$name, 's'=>$sz, 'p'=>$p, 't'=>time());
    room_save($c, $s);
    jout(array('ok'=>1, 'id'=>$id, 'count'=>count($s['files'])));
}

// ================= API: list =================
if ($api === 'list') {
    $c = clean_code(isset($_REQUEST['r']) ? $_REQUEST['r'] : '');
    $s = room_load($c);
    if ($s === null) jout(array('ok'=>0,'expired'=>1));
    $out = array();
    if (!empty($s['files'])) {
        foreach ($s['files'] as $fm)
            $out[] = array('id'=>$fm['id'], 'name'=>$fm['n'], 'size'=>intval($fm['s']));
    }
    jout(array('ok'=>1, 'files'=>$out,
               'left'=>max(0, intval($s['created']) + intval($s['ttl']) - time())));
}

// ================= API: download =================
if ($api === 'get') {
    $c = clean_code(isset($_REQUEST['r']) ? $_REQUEST['r'] : '');
    $id = substr(preg_replace('/[^a-f0-9]/', '', (string)(isset($_REQUEST['id']) ? $_REQUEST['id'] : '')), 0, 32);
    $s = room_load($c);
    if ($s === null) jout(array('ok'=>0,'err'=>'expired'), 410);
    $found = null;
    if (!empty($s['files'])) {
        foreach ($s['files'] as $fm) { if ($fm['id'] === $id) { $found = $fm; break; } }
    }
    if ($found === null) jout(array('ok'=>0,'err'=>'not found'), 404);
    $fp = $DIR . '/' . basename($found['p']);
    if (!is_file($fp)) jout(array('ok'=>0,'err'=>'not found'), 404);
    $dl = preg_replace('/[^A-Za-z0-9._ ()-]/', '_', (string)$found['n']);
    header('Content-Type: application/octet-stream');
    header('Content-Length: ' . filesize($fp));
    header('Content-Disposition: attachment; filename="' . $dl . '"');
    readfile($fp);
    exit;
}

// ================= WEB PAGE =================
$room = clean_code(isset($_GET['r']) ? $_GET['r'] : '');
$valid = ($room !== '' && room_load($room) !== null);
$self  = strtok($_SERVER['REQUEST_URI'], '?');      // path (bina query)
$host  = (isset($_SERVER['HTTP_HOST']) ? $_SERVER['HTTP_HOST'] : '');
$base  = 'https://' . $host . $self;
header('Content-Type: text/html; charset=utf-8');
?><!doctype html>
<html lang="hi"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>ApneScan Share — Phone to Phone</title>
<meta name="robots" content="noindex">
<style>
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
body{margin:0;font-family:system-ui,'Segoe UI',Roboto,Arial;background:#0b1220;color:#e6edf7;
  min-height:100vh;display:flex;flex-direction:column;align-items:center;padding:18px 14px 44px}
.card{width:100%;max-width:460px;background:#111a2e;border:1px solid #22304d;border-radius:20px;
  padding:20px 18px;box-shadow:0 18px 50px rgba(0,0,0,.45)}
h1{font-size:20px;margin:2px 0 4px;display:flex;align-items:center;gap:8px}
.sub{color:#93a4c3;font-size:12.5px;margin-bottom:14px;line-height:1.5}
.btn{display:flex;align-items:center;justify-content:center;gap:10px;width:100%;padding:15px 16px;margin:9px 0;
  border:none;border-radius:14px;background:#1b2946;color:#fff;font-size:15.5px;font-weight:700;cursor:pointer;text-align:center}
.btn:active{transform:scale(.985)}
.btn.pri{background:linear-gradient(135deg,#22b8a8,#0d9488)}
.btn.blue{background:linear-gradient(135deg,#2a78d6,#1f5fb0)}
.qrwrap{text-align:center;margin:10px 0 6px}
.qr{background:#fff;border-radius:14px;padding:10px;max-width:70%;height:auto}
.code{font-size:34px;font-weight:800;letter-spacing:6px;text-align:center;color:#5eead4;margin:6px 0 2px}
.linkbox{background:#0d1526;border:1px solid #22304d;border-radius:12px;padding:10px 12px;font-size:12px;
  color:#9db4e0;word-break:break-all;margin:8px 0}
.row{display:flex;gap:8px}.row .btn{margin:6px 0}
.item{background:#0d1526;border:1px solid #22304d;border-radius:12px;padding:11px 12px;margin:7px 0}
.item .nm{display:flex;justify-content:space-between;gap:8px;align-items:center}
.dl{background:none;border:1px solid #2f6f66;color:#5eead4;border-radius:9px;padding:5px 12px;font-size:12.5px;
  text-decoration:none;font-weight:700;white-space:nowrap}
.bar{height:8px;background:#1b2946;border-radius:6px;overflow:hidden;margin-top:7px}
.bar i{display:block;height:100%;width:0%;background:linear-gradient(90deg,#2a78d6,#22b8a8);transition:width .2s}
.st{font-weight:700}.ok{color:#3fd68f}.bad{color:#ff7d7d}
.exp{font-size:11.5px;color:#93a4c3;margin-top:10px;text-align:center}
.hint{font-size:12px;color:#93a4c3;margin:4px 2px 0}
.inp{width:100%;padding:13px;border-radius:12px;border:1px solid #2f4166;background:#0d1526;color:#fff;
  font-size:18px;letter-spacing:4px;text-align:center;text-transform:uppercase}
.foot{color:#5b6b8c;font-size:10.5px;margin-top:16px;text-align:center}
.dead{background:#2a1520;border:1px solid #5d2b3a;border-radius:14px;padding:16px;color:#ffb3c0;font-size:14px}
</style></head><body>
<div class="card">
<h1>📲 ApneScan Share</h1>

<?php if ($valid) { ?>
  <div class="sub">Ye room DONO taraf khula hai — jo bhi file yahan daaloge wo doosre device par turant dikhegi (aur ulta bhi). Koi bhi file chalegi.</div>
  <div class="qrwrap">
    <img class="qr" alt="QR" width="210" height="210"
         onerror="this.parentNode.style.display='none'"
         src="https://api.qrserver.com/v1/create-qr-code/?size=260x260&margin=0&data=<?php echo urlencode($base . '?r=' . $room); ?>">
    <div style="color:#93a4c3;font-size:12px;margin-top:6px">📷 Doosre phone se ye QR scan karo</div>
  </div>
  <div style="text-align:center;color:#93a4c3;font-size:12px">…ya code daalo</div>
  <div class="code"><?php echo htmlspecialchars($room); ?></div>
  <div class="linkbox" id="lnk"><?php echo htmlspecialchars($base . '?r=' . $room); ?></div>
  <div class="row">
    <button class="btn" style="flex:1" onclick="copyLink()">📋 Link copy</button>
    <a class="btn blue" style="flex:1;text-decoration:none" id="wa" href="#">🟢 WhatsApp</a>
  </div>
  <button class="btn pri" onclick="document.getElementById('f').click()">➕ File(s) attach karo</button>
  <input type="file" id="f" multiple style="display:none">
  <div id="list"></div>
  <div class="exp" id="exp"></div>
<?php } elseif ($room !== '') { ?>
  <div class="dead">⛔ Ye room expire ho gaya ya galat code hai.<br>Naya room banao ya sahi code daalo.</div>
  <button class="btn pri" onclick="location.href='<?php echo htmlspecialchars($self); ?>'">🏠 Naya room</button>
<?php } else { ?>
  <div class="sub">Bina kisi app ke — <b>phone se phone</b> (ya phone↔computer) koi bhi file bhejo.
    Ek device par room banao, doosre par wahi link/code kholo — bas.</div>
  <button class="btn pri" onclick="createRoom()">🚀 Naya Share Room banao</button>
  <div class="hint">…ya kisi ne code diya ho to yahan daalo:</div>
  <input class="inp" id="code" maxlength="6" placeholder="CODE" autocomplete="off">
  <button class="btn" onclick="joinRoom()">➡ Room me jao</button>
<?php } ?>
<div class="foot">ApneScan · Secure temporary room · files auto-delete</div>
</div>
<script>
var R = <?php echo json_encode($valid ? $room : ''); ?>;
var SELF = <?php echo json_encode($self); ?>;
function esc(s){return String(s).replace(/[&<>]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c];});}
function fmt(n){return n<1048576?Math.round(n/1024)+' KB':(n/1048576).toFixed(1)+' MB';}
function mmss(s){return Math.floor(s/60)+':'+('0'+(s%60)).slice(-2);}

function createRoom(){
  fetch('?api=create',{method:'POST'}).then(function(r){return r.json();}).then(function(j){
    if(j.ok){ location.href = SELF + '?r=' + j.code; } else { alert('Room nahi bana — dobara koshish karo'); }
  }).catch(function(){ alert('Network error'); });
}
function joinRoom(){
  var c=(document.getElementById('code').value||'').trim().toUpperCase();
  if(c.length>=4){ location.href = SELF + '?r=' + c; }
}
function copyLink(){
  var t=document.getElementById('lnk').textContent;
  (navigator.clipboard? navigator.clipboard.writeText(t): Promise.reject())
    .then(function(){ alert('Link copy ho gaya'); })
    .catch(function(){ prompt('Copy karo:', t); });
}

if (R) {
  var link = location.href.split('#')[0];
  document.getElementById('wa').href = 'https://wa.me/?text=' + encodeURIComponent('ApneScan Share — file bhejne/lene ke liye ye kholo: ' + link);

  document.getElementById('f').addEventListener('change', function(e){
    var fs=[].slice.call(e.target.files); e.target.value='';
    fs.forEach(function(file){ up(file); });
  });
  function up(file){
    var d=document.createElement('div'); d.className='item';
    d.innerHTML='<div class="nm"><span>'+esc(file.name)+'</span><span class="st">⏳</span></div><div class="bar"><i></i></div>';
    var L=document.getElementById('list'); L.insertBefore(d,L.firstChild);
    var st=d.querySelector('.st'), bar=d.querySelector('.bar i');
    var fd=new FormData(); fd.append('file',file,file.name);
    var x=new XMLHttpRequest(); x.open('POST','?api=up&r='+R);
    x.upload.onprogress=function(ev){ if(ev.total) bar.style.width=Math.round(ev.loaded*100/ev.total)+'%'; };
    x.onload=function(){ if(x.status===200){ st.className='st ok'; st.textContent='✅'; refresh(); }
      else { st.className='st bad'; st.textContent='❌ ' + ({410:'expire',413:'badi/full',429:'busy'}[x.status]||'fail'); } };
    x.onerror=function(){ st.className='st bad'; st.textContent='❌ net'; };
    x.send(fd);
  }
  function refresh(){
    var x=new XMLHttpRequest(); x.open('GET','?api=list&r='+R+'&_='+Date.now());
    x.onload=function(){ try{ var j=JSON.parse(x.responseText);
      if(!j.ok){ document.getElementById('exp').textContent='⛔ Room expire ho gaya'; return; }
      var h=''; (j.files||[]).forEach(function(f){
        h+='<div class="item"><div class="nm"><span>'+esc(f.name)+'</span>'+
           '<a class="dl" href="?api=get&r='+R+'&id='+encodeURIComponent(f.id)+'" download="'+esc(f.name)+'">⬇ Download</a></div>'+
           '<div style="color:#93a4c3;font-size:11px;margin-top:4px">'+fmt(f.size)+'</div></div>';
      });
      // upload-progress waale items ko mat mitao — sirf list refresh alag div me
      var L=document.getElementById('list');
      // purane server-list items hatao (jinke paas .dl hai) phir naye lagao upar-neeche
      [].slice.call(L.querySelectorAll('.item')).forEach(function(it){ if(it.querySelector('.dl')) it.remove(); });
      L.insertAdjacentHTML('beforeend', h);
      document.getElementById('exp').textContent='⏱ Room '+mmss(j.left)+' me expire hoga · files server se apne aap delete';
    }catch(e){} };
    x.send();
  }
  refresh(); setInterval(refresh, 4000);
}
</script></body></html>
