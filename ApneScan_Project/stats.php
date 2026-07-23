<?php
/**
 * ApneScan Stats — server + admin dashboard (Google Apps Script ki JAGAH).
 * ------------------------------------------------------------------------
 * Ek hi PHP file: (1) App se scan/ginti leti hai aur stats.json me store karti
 * hai (koi Google Sheet nahi), (2) App ko wahi JSON deti hai jo pehle Google
 * Script deta tha, (3) ?admin=PASSWORD par ek sundar dashboard dikhati hai.
 *
 * SETUP (ek baar):
 *   1. Is file ko apni hosting (Hostinger) par upload karo, jaise:
 *        https://apnesoftware.com/stats.php
 *   2. Neeche $ADMIN_PASS badal lo (admin panel ka password).
 *   3. App me Settings -> "Stats server URL" me yahi URL daal do.
 *   4. Admin panel: browser me kholo  https://apnesoftware.com/stats.php?admin=YOURPASS
 *
 * PRIVACY: yahan sirf GINTI aati hai. Kabhi koi document/patient data nahi.
 */

date_default_timezone_set('Asia/Kolkata');   // "Aaj" India (IST) ke hisaab se

$DATA_FILE  = __DIR__ . '/stats.json';        // data yahin store hoga
$ADMIN_PASS = 'apne123';                      // <-- ISE BADAL LO (admin password)
$SECRET     = '';                             // optional: scan protect karne ko

header('Access-Control-Allow-Origin: *');

// ---------------- storage helpers ----------------
function default_data() {
    return array('total'=>0,'days'=>array(),'hours'=>array(),'imports'=>0,'prints'=>0,
                 'peakAll'=>0,'peakDay'=>array(),'clients'=>array(),'online'=>array());
}
function load_data($file) {
    if (!file_exists($file)) return default_data();
    $d = json_decode(@file_get_contents($file), true);
    if (!is_array($d)) $d = default_data();
    return array_merge(default_data(), $d);
}
function save_data($file, $d) {
    $fp = @fopen($file, 'c+');
    if ($fp) {
        flock($fp, LOCK_EX); ftruncate($fp, 0); rewind($fp);
        fwrite($fp, json_encode($d)); fflush($fp); flock($fp, LOCK_UN); fclose($fp);
    }
}
function today_str() { return date('Y-m-d'); }
function hour_key()  { return date('Y-m-d-H'); }

function touch_client(&$d, $client, $req, $n, $now, $today) {
    if ($client === '') return;
    if (!isset($d['clients'][$client]))
        $d['clients'][$client] = array('first'=>$now,'last'=>$now,'version'=>'','country'=>'','scans'=>0,'method'=>'');
    $c =& $d['clients'][$client];
    $c['last'] = $now;
    if (!empty($req['v'])) $c['version'] = substr($req['v'], 0, 10);
    if (!empty($req['c'])) $c['country'] = substr($req['c'], 0, 4);
    if (!empty($req['m'])) $c['method']  = substr($req['m'], 0, 10);
    if ($n) $c['scans'] = intval($c['scans']) + $n;
    $d['online'][$client] = $now;
}
function update_peak(&$d, $now, $today) {
    $online = 0;
    foreach ($d['online'] as $ts) { if ($now - intval($ts) <= 300) $online++; }
    if ($online > intval($d['peakAll'])) $d['peakAll'] = $online;
    $pd = isset($d['peakDay'][$today]) ? intval($d['peakDay'][$today]) : 0;
    if ($online > $pd) $d['peakDay'][$today] = $online;
}

// ---------------- compute stats (App ko JSON) ----------------
function compute_stats($d, $client) {
    $today = today_str();
    $now = time();
    $online = 0;
    foreach ($d['online'] as $ts) { if ($now - intval($ts) <= 300) $online++; }

    $week = array(); $month = array(); $bestDay = 0;
    for ($j = 6; $j >= 0; $j--) { $k = date('Y-m-d', $now - $j*86400); $week[]  = array($k, intval(isset($d['days'][$k])?$d['days'][$k]:0)); }
    for ($j = 29; $j >= 0; $j--) { $k = date('Y-m-d', $now - $j*86400); $month[] = array($k, intval(isset($d['days'][$k])?$d['days'][$k]:0)); }
    foreach ($d['days'] as $v) { if (intval($v) > $bestDay) $bestDay = intval($v); }

    $versions = array(); $countries = array(); $methods = array(); $scores = array(); $mine = 0; $newToday = 0;
    foreach ($d['clients'] as $id => $c) {
        $v  = trim(isset($c['version'])?$c['version']:''); if ($v  !== '') $versions[$v]   = (isset($versions[$v])?$versions[$v]:0) + 1;
        $co = trim(isset($c['country'])?$c['country']:''); if ($co !== '') $countries[$co] = (isset($countries[$co])?$countries[$co]:0) + 1;
        $m  = trim(isset($c['method'])?$c['method']:'');  if ($m  !== '') $methods[$m]    = (isset($methods[$m])?$methods[$m]:0) + 1;
        $sc = intval(isset($c['scans'])?$c['scans']:0); $scores[] = $sc;
        if ($client !== '' && (string)$id === (string)$client) $mine = $sc;
        $fs = intval(isset($c['first'])?$c['first']:0);
        if ($fs && date('Y-m-d', $fs) === $today) $newToday++;
    }
    rsort($scores);
    $rank = 1; foreach ($scores as $s) { if ($s > $mine) $rank++; }
    $top = array_slice($scores, 0, 10);

    $todayHours = array_fill(0, 24, 0); $hourCount = 0; $curHK = hour_key();
    foreach ($d['hours'] as $hk => $hv) {
        if (strpos($hk, $today.'-') === 0) {
            $hh = intval(substr($hk, strlen($today)+1));
            if ($hh >= 0 && $hh < 24) $todayHours[$hh] = intval($hv);
        }
        if ($hk === $curHK) $hourCount = intval($hv);
    }

    return array(
        'ok'=>true, 'srv'=>'php1', 'time'=>date('Y-m-d H:i'),
        'today_key'=>'day_'.$today,
        'total'=>intval($d['total']),
        'today'=>intval(isset($d['days'][$today])?$d['days'][$today]:0),
        'online'=>$online, 'users'=>count($d['clients']), 'newToday'=>$newToday,
        'week'=>$week, 'month'=>$month, 'todayHours'=>$todayHours,
        'peak'=>intval(isset($d['peakDay'][$today])?$d['peakDay'][$today]:0),
        'peakAll'=>intval($d['peakAll']), 'bestDay'=>$bestDay, 'hour'=>$hourCount,
        'imports'=>intval($d['imports']), 'prints'=>intval($d['prints']),
        'versions'=>$versions, 'countries'=>$countries, 'methods'=>$methods,
        'top'=>$top, 'topscans'=>(count($top)?$top[0]:0), 'rank'=>$rank, 'myscans'=>$mine
    );
}

// ================= ADMIN — login + dashboard =================
// Clean URL:  apnescan.apnesoft.com/admin   (.htaccess -> stats.php?admin=1)
if (isset($_GET['admin'])) {
    session_start();
    // logout
    if (isset($_GET['logout'])) {
        $_SESSION = array(); @session_destroy();
        header('Location: ' . strtok($_SERVER['REQUEST_URI'], '?')); exit;
    }
    // login form POST
    $login_err = '';
    if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['pass'])) {
        if (hash_equals($ADMIN_PASS, (string)$_POST['pass'])) { $_SESSION['ok'] = true; }
        else { $login_err = 'Galat password. Dubara koshish karo.'; }
    }
    // ---- agar login nahi hua: login page dikhao ----
    if (empty($_SESSION['ok'])) {
        header('Content-Type: text/html; charset=utf-8');
        ?><!doctype html>
<html lang="hi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ApneScan Admin — Login</title>
<style>
  body{margin:0;font-family:system-ui,Segoe UI,Roboto,Arial;background:#0f172a;
    display:flex;align-items:center;justify-content:center;min-height:100vh;color:#e2e8f0}
  .box{background:#fff;color:#1e293b;width:320px;max-width:90%;padding:26px;border-radius:16px;
    box-shadow:0 12px 40px rgba(0,0,0,.35);text-align:center}
  .box .logo{font-size:34px} h1{font-size:19px;margin:6px 0 2px}
  .sub{color:#64748b;font-size:12px;margin-bottom:16px}
  input{width:100%;padding:11px 12px;border:1px solid #cbd5e1;border-radius:10px;font-size:15px;margin-bottom:10px}
  button{width:100%;padding:11px;border:none;border-radius:10px;background:#0f766e;color:#fff;
    font-size:15px;font-weight:700;cursor:pointer}
  button:hover{background:#0b5c55}
  .err{color:#dc2626;font-size:12px;margin-bottom:8px}
</style></head><body>
  <form class="box" method="post" action="">
    <div class="logo">📊🔒</div>
    <h1>ApneScan Admin</h1>
    <div class="sub">Worldwide stats dashboard</div>
    <?php if ($login_err) echo '<div class="err">'.htmlspecialchars($login_err).'</div>'; ?>
    <input type="password" name="pass" placeholder="Password" autofocus required>
    <button type="submit">Login →</button>
  </form>
</body></html><?php
        exit;
    }
    // ---- login ho gaya: dashboard ----
    $d = load_data($DATA_FILE);
    $S = compute_stats($d, '');
    $J = json_encode($S);
    ?><!doctype html>
<html lang="hi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ApneScan — Admin Stats</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root{--te:#0f766e;--bg:#f6f8fa;--card:#fff;--mut:#64748b;--line:#e2e8f0}
  *{box-sizing:border-box} body{margin:0;font-family:system-ui,Segoe UI,Roboto,Arial;background:var(--bg);color:#1e293b}
  header{background:linear-gradient(90deg,#0f766e,#0891b2);color:#fff;padding:16px 22px}
  header h1{margin:0;font-size:20px} header .t{opacity:.85;font-size:12px;margin-top:2px}
  .wrap{max-width:1100px;margin:18px auto;padding:0 14px}
  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:16px}
  .kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px}
  .kpi .n{font-size:26px;font-weight:800;color:var(--te)} .kpi .l{color:var(--mut);font-size:12px;margin-top:2px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
  @media(max-width:760px){.grid{grid-template-columns:1fr}}
  .card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px}
  .card h3{margin:0 0 10px;font-size:14px}
  table{width:100%;border-collapse:collapse;font-size:13px} td{padding:5px 4px;border-bottom:1px solid var(--line)}
  .bar{height:10px;background:var(--te);border-radius:5px}
  .foot{color:var(--mut);font-size:12px;text-align:center;margin:18px 0}
  .rbtn{float:right;background:#fff;color:#0f766e;border:1px solid #fff;border-radius:8px;padding:5px 12px;cursor:pointer;font-weight:700}
</style></head><body>
<header>
  <a class="rbtn" href="?logout=1" style="text-decoration:none;margin-left:8px">🔓 Logout</a>
  <button class="rbtn" onclick="location.reload()">🔄 Refresh</button>
  <h1>📊 ApneScan — Worldwide Stats</h1>
  <div class="t">Live data · <span id="tm"></span> · sirf ginti (koi document/naam nahi)</div>
</header>
<div class="wrap">
  <div class="kpis" id="kpis"></div>
  <div class="grid">
    <div class="card"><h3>📊 Last 7 days</h3><canvas id="wk" height="150"></canvas></div>
    <div class="card"><h3>📈 Last 30 days</h3><canvas id="mo" height="150"></canvas></div>
    <div class="card"><h3>🌍 Today — 24 hours</h3><canvas id="hr" height="150"></canvas></div>
    <div class="card"><h3>🏆 Top users (anonymous)</h3><canvas id="tp" height="150"></canvas></div>
    <div class="card"><h3>🗺 Countries</h3><div id="co"></div></div>
    <div class="card"><h3>🔢 App versions</h3><div id="ve"></div></div>
    <div class="card"><h3>🖨 Scan methods</h3><div id="me"></div></div>
    <div class="card"><h3>🏅 Records</h3><div id="rc"></div></div>
  </div>
  <div class="foot">Server: PHP · <?php echo htmlspecialchars($S['time']); ?> · ApneSoftware.com</div>
</div>
<script>
var D = <?php echo $J; ?>;
document.getElementById('tm').textContent = D.time;
function kpi(n,l){return '<div class="kpi"><div class="n">'+n+'</div><div class="l">'+l+'</div></div>';}
document.getElementById('kpis').innerHTML =
  kpi(D.total.toLocaleString(),'Total scans')+kpi(D.today,'Today')+kpi(D.online,'Online now')+
  kpi(D.users,'Users')+kpi(D.newToday,'New today')+kpi(D.imports,'Imports')+kpi(D.prints,'Prints');
function bars(id,rows){var h='<table>';rows.forEach(function(r){var mx=rows[0]?rows[0][1]:1;
  h+='<tr><td style="width:90px">'+r[0]+'</td><td><div class="bar" style="width:'+Math.max(4,100*r[1]/(mx||1))+'%"></div></td><td style="width:40px;text-align:right">'+r[1]+'</td></tr>';});
  document.getElementById(id).innerHTML=h+'</table>';}
function obj2rows(o){return Object.keys(o).map(function(k){return [k,o[k]];}).sort(function(a,b){return b[1]-a[1];}).slice(0,8);}
new Chart(wk,{type:'bar',data:{labels:D.week.map(function(x){return x[0].slice(5);}),datasets:[{data:D.week.map(function(x){return x[1];}),backgroundColor:'#0f766e'}]},options:{plugins:{legend:{display:false}}}});
new Chart(mo,{type:'line',data:{labels:D.month.map(function(x){return x[0].slice(5);}),datasets:[{data:D.month.map(function(x){return x[1];}),borderColor:'#0891b2',backgroundColor:'rgba(8,145,178,.15)',fill:true,tension:.3,pointRadius:0}]},options:{plugins:{legend:{display:false}}}});
new Chart(hr,{type:'bar',data:{labels:D.todayHours.map(function(_,i){return i;}),datasets:[{data:D.todayHours,backgroundColor:'#0f766e'}]},options:{plugins:{legend:{display:false}}}});
new Chart(tp,{type:'bar',data:{labels:D.top.map(function(_,i){return 'User '+(i+1);}),datasets:[{data:D.top,backgroundColor:'#7c3aed'}]},options:{indexAxis:'y',plugins:{legend:{display:false}}}});
bars('co',obj2rows(D.countries));bars('ve',obj2rows(D.versions));bars('me',obj2rows(D.methods));
document.getElementById('rc').innerHTML='<table>'+
 '<tr><td>🏔 Peak online (all-time)</td><td style="text-align:right"><b>'+D.peakAll+'</b></td></tr>'+
 '<tr><td>📅 Best single day</td><td style="text-align:right"><b>'+D.bestDay+'</b></td></tr>'+
 '<tr><td>🕐 This hour</td><td style="text-align:right"><b>'+D.hour+'</b></td></tr>'+
 '<tr><td>🖥 Server</td><td style="text-align:right"><b>'+D.srv+'</b></td></tr></table>';
</script>
</body></html><?php
    exit;
}

// ================= API (App se) =================
$action = isset($_REQUEST['action']) ? $_REQUEST['action'] : 'stats';
$client = isset($_REQUEST['client']) ? $_REQUEST['client'] : '';
$today  = today_str();
$now    = time();

$d = load_data($DATA_FILE);

if ($action === 'scan') {
    if ($SECRET !== '' && (!isset($_REQUEST['secret']) || $_REQUEST['secret'] !== $SECRET)) {
        header('Content-Type: application/json');
        echo json_encode(array('ok'=>false,'error'=>'bad secret')); exit;
    }
    $n = max(0, min(100, intval(isset($_REQUEST['n'])?$_REQUEST['n']:1)));
    $d['total'] = intval($d['total']) + $n;
    $d['days'][$today] = (isset($d['days'][$today])?intval($d['days'][$today]):0) + $n;
    $hk = hour_key();
    $d['hours'][$hk] = (isset($d['hours'][$hk])?intval($d['hours'][$hk]):0) + $n;
    touch_client($d, $client, $_REQUEST, $n, $now, $today);
    update_peak($d, $now, $today);
} else if ($action === 'ping') {
    touch_client($d, $client, $_REQUEST, 0, $now, $today);
    update_peak($d, $now, $today);
}
// import/print kisi bhi action ke saath
$imp = max(0, min(500, intval(isset($_REQUEST['imp'])?$_REQUEST['imp']:0)));
$prt = max(0, min(500, intval(isset($_REQUEST['prt'])?$_REQUEST['prt']:0)));
if ($imp) $d['imports'] = intval($d['imports']) + $imp;
if ($prt) $d['prints']  = intval($d['prints'])  + $prt;

// online list ko chhota rakho (purane 1 din se zyada purane hata do)
foreach ($d['online'] as $id => $ts) { if ($now - intval($ts) > 86400) unset($d['online'][$id]); }

save_data($DATA_FILE, $d);

header('Content-Type: application/json');
echo json_encode(compute_stats($d, $client));
