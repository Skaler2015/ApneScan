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
        $d['clients'][$client] = array('first'=>$now,'last'=>$now,'version'=>'','country'=>'','scans'=>0,'method'=>'','name'=>'');
    $c =& $d['clients'][$client];
    $c['last'] = $now;
    if (!empty($req['v'])) $c['version'] = substr($req['v'], 0, 10);
    if (!empty($req['c'])) $c['country'] = substr($req['c'], 0, 4);
    if (!empty($req['m'])) $c['method']  = substr($req['m'], 0, 10);
    if (!empty($req['u'])) $c['name']    = substr($req['u'], 0, 40);   // user ka naam
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

    $versions = array(); $countries = array(); $methods = array(); $scores = array();
    $named = array(); $mine = 0; $newToday = 0;
    foreach ($d['clients'] as $id => $c) {
        $v  = trim(isset($c['version'])?$c['version']:''); if ($v  !== '') $versions[$v]   = (isset($versions[$v])?$versions[$v]:0) + 1;
        $co = trim(isset($c['country'])?$c['country']:''); if ($co !== '') $countries[$co] = (isset($countries[$co])?$countries[$co]:0) + 1;
        $m  = trim(isset($c['method'])?$c['method']:'');  if ($m  !== '') $methods[$m]    = (isset($methods[$m])?$methods[$m]:0) + 1;
        $sc = intval(isset($c['scans'])?$c['scans']:0); $scores[] = $sc;
        $nm = trim(isset($c['name'])?$c['name']:'');
        $named[] = array('name'=>($nm !== '' ? $nm : '—'), 'scans'=>$sc);
        if ($client !== '' && (string)$id === (string)$client) $mine = $sc;
        $fs = intval(isset($c['first'])?$c['first']:0);
        if ($fs && date('Y-m-d', $fs) === $today) $newToday++;
    }
    rsort($scores);
    $rank = 1; foreach ($scores as $s) { if ($s > $mine) $rank++; }
    $top = array_slice($scores, 0, 10);
    usort($named, function($a, $b) { return $b['scans'] - $a['scans']; });
    $topNamed = array_slice($named, 0, 10);

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
        'top'=>$top, 'topNamed'=>$topNamed, 'topscans'=>(count($top)?$top[0]:0),
        'rank'=>$rank, 'myscans'=>$mine
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
    // ---- login ho gaya ----
    $t0 = microtime(true);
    $d = load_data($DATA_FILE);

    // login log (session me ek baar)
    if (empty($_SESSION['logged'])) {
        $_SESSION['logged'] = true;
        if (!isset($d['adminLogins'])) $d['adminLogins'] = array();
        $d['adminLogins'][] = time();
        $d['adminLogins'] = array_slice($d['adminLogins'], -15);
        save_data($DATA_FILE, $d);
    }

    // ---- EXPORT: CSV / JSON backup ----
    if (isset($_GET['export'])) {
        $e = $_GET['export'];
        if ($e === 'json') {
            header('Content-Type: application/json');
            header('Content-Disposition: attachment; filename="apnescan-stats-backup.json"');
            echo json_encode($d); exit;
        }
        if ($e === 'days') {
            header('Content-Type: text/csv; charset=utf-8');
            header('Content-Disposition: attachment; filename="apnescan-daily.csv"');
            echo "date,scans\n"; ksort($d['days']);
            foreach ($d['days'] as $dt => $c) echo $dt . "," . intval($c) . "\n";
            exit;
        }
        if ($e === 'users') {
            header('Content-Type: text/csv; charset=utf-8');
            header('Content-Disposition: attachment; filename="apnescan-users.csv"');
            echo "name,scans,first_seen,last_seen,version,country,method\n";
            foreach ($d['clients'] as $c) {
                echo '"' . str_replace('"', '""', trim(isset($c['name'])?$c['name']:'')) . '",'
                    . intval(isset($c['scans'])?$c['scans']:0) . ','
                    . date('Y-m-d H:i', intval(isset($c['first'])?$c['first']:0)) . ','
                    . date('Y-m-d H:i', intval(isset($c['last'])?$c['last']:0)) . ','
                    . trim(isset($c['version'])?$c['version']:'') . ','
                    . trim(isset($c['country'])?$c['country']:'') . ','
                    . trim(isset($c['method'])?$c['method']:'') . "\n";
            }
            exit;
        }
    }

    // ---- rich data compute ----
    $S = compute_stats($d, '');
    $today = today_str(); $now = time();
    $yest = date('Y-m-d', $now - 86400);
    $S['yesterday'] = intval(isset($d['days'][$yest])?$d['days'][$yest]:0);
    $wt = 0; for ($i=0;$i<7;$i++)  { $k=date('Y-m-d',$now-$i*86400); $wt += intval(isset($d['days'][$k])?$d['days'][$k]:0); }
    $mt = 0; for ($i=0;$i<30;$i++) { $k=date('Y-m-d',$now-$i*86400); $mt += intval(isset($d['days'][$k])?$d['days'][$k]:0); }
    $S['weekTotal']=$wt; $S['monthTotal']=$mt; $S['dailyAvg']=round($mt/30,1);
    // DAU/WAU/MAU + joins
    $dau=$wau=$mau=0; $joinByDay=array();
    foreach ($d['clients'] as $c) {
        $last = intval(isset($c['last'])?$c['last']:0);
        if (date('Y-m-d',$last) === $today) $dau++;
        if ($now-$last <= 7*86400) $wau++;
        if ($now-$last <= 30*86400) $mau++;
        $fd = date('Y-m-d', intval(isset($c['first'])?$c['first']:0));
        $joinByDay[$fd] = (isset($joinByDay[$fd])?$joinByDay[$fd]:0) + 1;
    }
    $S['dau']=$dau; $S['wau']=$wau; $S['mau']=$mau;
    $newU=array(); $cumU=array();
    for ($i=29;$i>=0;$i--) {
        $k = date('Y-m-d',$now-$i*86400);
        $newU[] = array($k, isset($joinByDay[$k])?$joinByDay[$k]:0);
        $end = strtotime($k.' 23:59:59'); $cc=0;
        foreach ($d['clients'] as $c) { if (intval(isset($c['first'])?$c['first']:0) <= $end) $cc++; }
        $cumU[] = array($k, $cc);
    }
    $S['newUsers30']=$newU; $S['cumUsers30']=$cumU;
    // scans by country
    $scByCo=array();
    foreach ($d['clients'] as $c) { $co=trim(isset($c['country'])?$c['country']:''); if($co!=='') $scByCo[$co]=(isset($scByCo[$co])?$scByCo[$co]:0)+intval(isset($c['scans'])?$c['scans']:0); }
    $S['scansByCountry']=$scByCo;
    // user list + online names
    $userList=array(); $onlineNames=array();
    foreach ($d['clients'] as $id => $c) {
        $nm=trim(isset($c['name'])?$c['name']:''); $last=intval(isset($c['last'])?$c['last']:0);
        $userList[] = array('name'=>($nm!==''?$nm:'—'),'scans'=>intval(isset($c['scans'])?$c['scans']:0),
            'first'=>intval(isset($c['first'])?$c['first']:0),'last'=>$last,
            'version'=>trim(isset($c['version'])?$c['version']:''),'country'=>trim(isset($c['country'])?$c['country']:''),
            'method'=>trim(isset($c['method'])?$c['method']:''),'online'=>(($now-$last)<=300));
        if (($now-$last)<=300 && $nm!=='') $onlineNames[]=$nm;
    }
    usort($userList, function($a,$b){ return $b['scans']-$a['scans']; });
    $S['userList']=$userList; $S['onlineNames']=$onlineNames;
    // old version %
    $vers=$S['versions']; $tv=array_sum($vers); $latest=0; $oldpct=0;
    if ($tv>0) { foreach($vers as $k=>$v){ if(is_numeric($k)&&intval($k)>$latest)$latest=intval($k);} $old=0; foreach($vers as $k=>$v){ if(intval($k)!==$latest)$old+=$v;} $oldpct=round(100*$old/$tv); }
    $S['oldVersionPct']=$oldpct; $S['latestVersion']=$latest;
    $S['unusual'] = ($S['dailyAvg']>=3 && $S['today'] >= 3*$S['dailyAvg']);
    $S['adminLogins'] = array_map(function($ts){ return date('Y-m-d H:i',$ts); },
        isset($d['adminLogins']) ? array_reverse(array_slice($d['adminLogins'],-8)) : array());
    $S['daysMap'] = $d['days'];
    $S['respMs'] = round((microtime(true)-$t0)*1000);

    $J = json_encode($S);
    ?><!doctype html>
<html lang="hi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ApneScan — Admin Stats</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root{--te:#0f766e;--bg:#f6f8fa;--card:#fff;--mut:#64748b;--line:#e2e8f0}
  *{box-sizing:border-box} body{margin:0;font-family:system-ui,Segoe UI,Roboto,Arial;background:var(--bg);color:#1e293b}
  header{background:linear-gradient(90deg,#0f766e,#0891b2);color:#fff;padding:14px 20px;position:sticky;top:0;z-index:5}
  header h1{margin:0;font-size:19px} header .t{opacity:.85;font-size:12px;margin-top:2px}
  .wrap{max-width:1150px;margin:16px auto;padding:0 14px}
  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin-bottom:14px}
  .kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:12px}
  .kpi .n{font-size:23px;font-weight:800;color:var(--te)} .kpi .l{color:var(--mut);font-size:11px;margin-top:2px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
  @media(max-width:760px){.grid{grid-template-columns:1fr}}
  .card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px;margin-bottom:14px}
  .card h3{margin:0 0 10px;font-size:14px}
  table{width:100%;border-collapse:collapse;font-size:13px} td,th{padding:5px 6px;border-bottom:1px solid var(--line);text-align:left}
  th{cursor:pointer;color:#334155;font-size:12px;background:#f8fafc}
  .bar{height:10px;background:var(--te);border-radius:5px}
  .rbtn{background:#fff;color:#0f766e;border:none;border-radius:8px;padding:6px 12px;cursor:pointer;font-weight:700;margin-left:6px;text-decoration:none;display:inline-block}
  .toolbar{float:right}
  input,select{padding:7px 9px;border:1px solid #cbd5e1;border-radius:8px;font-size:13px}
  .pill{display:inline-block;background:#dcfce7;color:#166534;border-radius:20px;padding:1px 8px;font-size:11px}
  .pill.off{background:#f1f5f9;color:#94a3b8}
  .banner{background:#fef9c3;border:1px solid #fde68a;border-radius:10px;padding:8px 12px;margin-bottom:12px;font-size:13px;display:none}
  .btns{margin-bottom:12px} .foot{color:var(--mut);font-size:12px;text-align:center;margin:16px 0}
  @media print{.toolbar,.btns,.no-print{display:none}}
</style></head><body>
<header>
  <div class="toolbar no-print">
    <button class="rbtn" onclick="location.reload()">🔄 Refresh</button>
    <button class="rbtn" onclick="window.print()">🖨 Print</button>
    <a class="rbtn" href="?logout=1">🔓 Logout</a>
  </div>
  <h1>📊 ApneScan — Admin Panel</h1>
  <div class="t">Live · <span id="tm"></span> · sirf ginti (koi document/naam nahi) · <span id="rt"></span></div>
</header>
<div class="wrap">
  <div id="banner" class="banner"></div>
  <div class="kpis" id="kpis"></div>

  <div class="btns no-print">
    📥 <a class="rbtn" href="?admin=1&export=days">Daily CSV</a>
    <a class="rbtn" href="?admin=1&export=users">Users CSV</a>
    <a class="rbtn" href="?admin=1&export=json">Backup (JSON)</a>
  </div>

  <div class="card no-print"><h3>🔍 Kisi bhi range ka jod</h3>
    <input type="date" id="df"> → <input type="date" id="dt">
    <b id="drsum" style="margin-left:10px"></b>
  </div>

  <div class="grid">
    <div class="card"><h3>📊 Last 7 days</h3><canvas id="wk" height="150"></canvas></div>
    <div class="card"><h3>📈 Last 30 days</h3><canvas id="mo" height="150"></canvas></div>
    <div class="card"><h3>🌍 Today — 24 hours</h3><canvas id="hr" height="150"></canvas></div>
    <div class="card"><h3>🏆 Top users (naam ke saath)</h3><canvas id="tp" height="150"></canvas></div>
    <div class="card"><h3>🌱 New users — 30 days</h3><canvas id="nu" height="150"></canvas></div>
    <div class="card"><h3>👥 Total users (growth)</h3><canvas id="cu" height="150"></canvas></div>
  </div>

  <div class="card"><h3>👤 Saare users (<span id="ucount"></span>) — click header to sort</h3>
    <input id="usearch" class="no-print" placeholder="🔍 naam / desh / version se dhoondo…" style="width:100%;margin-bottom:8px">
    <div style="overflow:auto;max-height:460px"><table id="utable"></table></div>
  </div>

  <div class="grid">
    <div class="card"><h3>🟢 Abhi online (<span id="oncount"></span>)</h3><div id="onlist"></div></div>
    <div class="card"><h3>⏱ Recently active</h3><div id="recent"></div></div>
    <div class="card"><h3>🗺 Countries (users)</h3><div id="co"></div></div>
    <div class="card"><h3>🗺 Scans by country</h3><div id="cos"></div></div>
    <div class="card"><h3>🔢 App versions</h3><div id="ve"></div></div>
    <div class="card"><h3>🖨 Scan methods</h3><div id="me"></div></div>
    <div class="card"><h3>🏅 Records</h3><div id="rc"></div></div>
    <div class="card"><h3>🔒 Admin logins</h3><div id="al"></div></div>
  </div>

  <div class="foot">Server: PHP <?php echo htmlspecialchars($S['srv']); ?> · <?php echo htmlspecialchars($S['time']); ?> · ApneSoftware.com</div>
</div>
<script>
var D = <?php echo $J; ?>;
document.getElementById('tm').textContent = D.time;
document.getElementById('rt').textContent = 'server '+D.respMs+'ms';
function flag(cc){ if(!cc||cc.length!==2) return '🏳'; return String.fromCodePoint.apply(null,[].map.call(cc.toUpperCase(),function(ch){return 127397+ch.charCodeAt(0);})); }
function fmt(n){ return (n||0).toLocaleString(); }
function ago(ts){ if(!ts)return '—'; var s=Math.floor(Date.now()/1000)-ts; if(s<60)return s+'s'; if(s<3600)return Math.floor(s/60)+'m'; if(s<86400)return Math.floor(s/3600)+'h'; return Math.floor(s/86400)+'d'; }

// KPIs
function kpi(n,l){return '<div class="kpi"><div class="n">'+n+'</div><div class="l">'+l+'</div></div>';}
document.getElementById('kpis').innerHTML =
  kpi(fmt(D.total),'Total scans')+kpi(D.today,'Today')+kpi(D.yesterday,'Yesterday')+
  kpi(D.weekTotal,'This week')+kpi(D.monthTotal,'This month')+kpi(D.dailyAvg,'Daily avg')+
  kpi(D.online,'Online now')+kpi(D.users,'Total users')+kpi(D.newToday,'New today')+
  kpi(D.dau,'Active today')+kpi(D.wau,'Active 7d')+kpi(D.mau,'Active 30d')+
  kpi(D.imports,'Imports')+kpi(D.prints,'Prints')+kpi(fmt(D.total),'Paper saved');

// banner
var bmsg='';
if(D.unusual) bmsg='🚀 Aaj normal se bahut zyada activity ho rahi hai!';
[100000,50000,10000,5000,1000,500,100].some(function(m){ if(D.total>=m && D.total-D.today<m){ bmsg='🎉 Milestone! Duniya bhar me '+fmt(m)+' scans paar!'; return true;} return false;});
if(D.oldVersionPct>=30) bmsg=(bmsg?bmsg+'  ':'')+'🔔 '+D.oldVersionPct+'% users purane version par hain.';
if(bmsg){ var b=document.getElementById('banner'); b.textContent=bmsg; b.style.display='block'; }

// bars helper
function bars(id,rows,useFlag){var mx=rows[0]?rows[0][1]:1;var h='<table>';rows.forEach(function(r){
  var lab=useFlag?(flag(r[0])+' '+r[0]):r[0];
  h+='<tr><td style="width:110px">'+lab+'</td><td><div class="bar" style="width:'+Math.max(4,100*r[1]/(mx||1))+'%"></div></td><td style="width:44px;text-align:right">'+r[1]+'</td></tr>';});
  document.getElementById(id).innerHTML=h+'</table>';}
function obj2rows(o){return Object.keys(o).map(function(k){return [k,o[k]];}).sort(function(a,b){return b[1]-a[1];}).slice(0,12);}

// charts
new Chart(wk,{type:'bar',data:{labels:D.week.map(function(x){return x[0].slice(5);}),datasets:[{data:D.week.map(function(x){return x[1];}),backgroundColor:'#0f766e'}]},options:{plugins:{legend:{display:false}}}});
new Chart(mo,{type:'line',data:{labels:D.month.map(function(x){return x[0].slice(5);}),datasets:[{data:D.month.map(function(x){return x[1];}),borderColor:'#0891b2',backgroundColor:'rgba(8,145,178,.15)',fill:true,tension:.3,pointRadius:0}]},options:{plugins:{legend:{display:false}}}});
new Chart(hr,{type:'bar',data:{labels:D.todayHours.map(function(_,i){return i;}),datasets:[{data:D.todayHours,backgroundColor:'#0f766e'}]},options:{plugins:{legend:{display:false}}}});
var TN=(D.topNamed&&D.topNamed.length)?D.topNamed:(D.top||[]).map(function(s,i){return {name:'User '+(i+1),scans:s};});
new Chart(tp,{type:'bar',data:{labels:TN.map(function(t,i){return (t.name&&t.name!=='—')?t.name:('User '+(i+1));}),datasets:[{data:TN.map(function(t){return t.scans;}),backgroundColor:'#7c3aed'}]},options:{indexAxis:'y',plugins:{legend:{display:false}}}});
new Chart(nu,{type:'bar',data:{labels:D.newUsers30.map(function(x){return x[0].slice(5);}),datasets:[{data:D.newUsers30.map(function(x){return x[1];}),backgroundColor:'#16a34a'}]},options:{plugins:{legend:{display:false}}}});
new Chart(cu,{type:'line',data:{labels:D.cumUsers30.map(function(x){return x[0].slice(5);}),datasets:[{data:D.cumUsers30.map(function(x){return x[1];}),borderColor:'#7c3aed',backgroundColor:'rgba(124,58,237,.12)',fill:true,tension:.3,pointRadius:0}]},options:{plugins:{legend:{display:false}}}});

// breakdowns
bars('co',obj2rows(D.countries),true);
bars('cos',obj2rows(D.scansByCountry),true);
bars('ve',obj2rows(D.versions),false);
bars('me',Object.keys(D.methods||{}).map(function(k){var n={escl:'Network',wia:'USB',twain:'TWAIN',naps2:'NAPS2'}[k]||k;return [n,D.methods[k]];}).sort(function(a,b){return b[1]-a[1];}),false);

// online + recent
document.getElementById('oncount').textContent=(D.onlineNames||[]).length;
document.getElementById('onlist').innerHTML=(D.onlineNames&&D.onlineNames.length)?D.onlineNames.map(function(n){return '<div>🟢 '+n+'</div>';}).join(''):'<div style="color:#94a3b8">Abhi koi online nahi</div>';
var rec=(D.userList||[]).slice().sort(function(a,b){return b.last-a.last;}).slice(0,10);
document.getElementById('recent').innerHTML='<table>'+rec.map(function(u){return '<tr><td>'+(u.online?'🟢 ':'⚪ ')+u.name+'</td><td style="text-align:right;color:#94a3b8">'+ago(u.last)+' ago</td></tr>';}).join('')+'</table>';

// records
document.getElementById('rc').innerHTML='<table>'+
 '<tr><td>🏔 Peak online (all-time)</td><td style="text-align:right"><b>'+D.peakAll+'</b></td></tr>'+
 '<tr><td>📅 Best single day</td><td style="text-align:right"><b>'+D.bestDay+'</b></td></tr>'+
 '<tr><td>🕐 This hour</td><td style="text-align:right"><b>'+D.hour+'</b></td></tr>'+
 '<tr><td>🆕 Latest version</td><td style="text-align:right"><b>v'+D.latestVersion+'</b></td></tr></table>';
document.getElementById('al').innerHTML='<table>'+(D.adminLogins||[]).map(function(t){return '<tr><td>🔒 '+t+'</td></tr>';}).join('')+'</table>';

// ---- user table: search + sort ----
var sortKey='scans', sortDir=-1;
function renderUsers(){
  var q=(document.getElementById('usearch').value||'').toLowerCase();
  var rows=(D.userList||[]).filter(function(u){ return (u.name+' '+u.country+' '+u.version+' '+u.method).toLowerCase().indexOf(q)>=0; });
  rows.sort(function(a,b){ var x=a[sortKey],y=b[sortKey]; if(typeof x==='string'){return sortDir*x.localeCompare(y);} return sortDir*((x||0)-(y||0)); });
  document.getElementById('ucount').textContent=rows.length;
  var h='<tr>'+[['name','Name'],['scans','Scans'],['last','Last seen'],['first','Joined'],['version','Ver'],['country','Desh'],['method','Method']].map(function(c){return '<th data-k="'+c[0]+'">'+c[1]+(sortKey===c[0]?(sortDir<0?' ▼':' ▲'):'')+'</th>';}).join('')+'</tr>';
  rows.forEach(function(u){ h+='<tr><td>'+(u.online?'🟢 ':'')+u.name+'</td><td><b>'+u.scans+'</b></td><td>'+ago(u.last)+' ago</td><td>'+(u.first?new Date(u.first*1000).toISOString().slice(0,10):'—')+'</td><td>'+(u.version||'—')+'</td><td>'+(u.country?flag(u.country)+' '+u.country:'—')+'</td><td>'+({escl:'Network',wia:'USB',twain:'TWAIN',naps2:'NAPS2'}[u.method]||u.method||'—')+'</td></tr>'; });
  var t=document.getElementById('utable'); t.innerHTML=h;
  [].forEach.call(t.querySelectorAll('th'),function(th){ th.onclick=function(){ var k=th.getAttribute('data-k'); if(sortKey===k)sortDir*=-1; else {sortKey=k;sortDir=(k==='name'||k==='country'||k==='version'||k==='method')?1:-1;} renderUsers(); }; });
}
document.getElementById('usearch').addEventListener('input',renderUsers);
renderUsers();

// ---- date-range sum ----
function drCalc(){
  var a=document.getElementById('df').value, b=document.getElementById('dt').value;
  if(!a||!b){document.getElementById('drsum').textContent='';return;}
  var s=0; for(var k in D.daysMap){ if(k>=a&&k<=b) s+=parseInt(D.daysMap[k])||0; }
  document.getElementById('drsum').textContent=fmt(s)+' scans';
}
(function(){var t=new Date(),f=new Date(t.getTime()-6*86400000);
 document.getElementById('dt').value=t.toISOString().slice(0,10);
 document.getElementById('df').value=f.toISOString().slice(0,10);})();
document.getElementById('df').addEventListener('change',drCalc);
document.getElementById('dt').addEventListener('change',drCalc); drCalc();

// LIVE auto-refresh (30s)
setInterval(function(){ location.reload(); }, 30000);
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
