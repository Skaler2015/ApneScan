<?php
/**
 * ApneScan Stats — server + FULL admin dashboard (Google Apps Script ki JAGAH).
 * ---------------------------------------------------------------------------
 * Ek hi PHP file:
 *   (1) App se scan/ginti/feature/crash/feedback leti hai -> stats.json me store
 *   (2) App ko wahi JSON deti hai (worldwide numbers + broadcast + remote-config)
 *   (3) ?admin=PASSWORD par ek poora dashboard dikhati hai
 *   (4) ?cron=daily&key=CRONKEY par roz email report + auto-backup
 *
 * SETUP (ek baar):
 *   1. File ko Hostinger par upload karo (jaise status.apnesoft.com/stats.php)
 *   2. Neeche $ADMIN_PASS, $ADMIN_EMAIL, $CRON_KEY badlo
 *   3. App me Settings -> "Stats server URL" me yahi URL
 *   4. Admin: browser me  .../stats.php?admin=YOURPASS   (ya  .../admin)
 *   5. (optional) Hostinger cron: roz 8am ->  php .../stats.php cron=daily key=CRONKEY
 *
 * PRIVACY: sirf GINTI/naam aata hai. Kabhi koi document/patient data nahi.
 */

date_default_timezone_set('Asia/Kolkata');   // "Aaj" India (IST) ke hisaab se

$DATA_FILE   = __DIR__ . '/stats.json';
$BACKUP_DIR  = __DIR__ . '/backups';
$ADMIN_PASS  = 'apne123';            // <-- ISE BADAL LO (admin password)
$ADMIN_EMAIL = '';                   // <-- daily report yahan aayega (khaali = band)
$CRON_KEY    = 'apnecron';           // <-- ?cron=daily&key=... ka key
$SECRET      = '';                   // optional: scan protect karne ko
$SESSION_TTL = 1800;                 // admin auto-logout (sec) — 30 min
$MAX_FAILS   = 6;                    // itni galat koshish -> thodi der lock

header('Access-Control-Allow-Origin: *');

// ---------------- storage helpers ----------------
function default_data() {
    return array(
        'total'=>0,'days'=>array(),'hours'=>array(),'imports'=>0,'prints'=>0,
        'peakAll'=>0,'peakDay'=>array(),'clients'=>array(),'online'=>array(),
        'features'=>array(),'scanners'=>array(),'dpis'=>array(),
        'colors'=>array(),'sizes'=>array(),'crashes'=>array(),'feedback'=>array(),
        'broadcast'=>array('msg'=>'','target'=>'all','id'=>0),'rconfig'=>array(),
        'adminLogins'=>array(),'failLog'=>array(),'lastBackup'=>0
    );
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
function bump(&$arr, $key, $by=1) { $key=trim((string)$key); if($key==='')return; $arr[$key]=(isset($arr[$key])?intval($arr[$key]):0)+$by; }

function touch_client(&$d, $client, $req, $n, $now, $today) {
    if ($client === '') return true;
    if (!isset($d['clients'][$client]))
        $d['clients'][$client] = array('first'=>$now,'last'=>$now,'version'=>'','country'=>'',
            'scans'=>0,'method'=>'','name'=>'','model'=>'','note'=>'','tags'=>'','blocked'=>0,
            'days'=>array(),'active'=>array());
    $c =& $d['clients'][$client];
    if (!empty($c['blocked'])) return false;              // blocked user -> ignore
    $c['last'] = $now;
    if (!empty($req['v']))  $c['version'] = substr($req['v'], 0, 10);
    if (!empty($req['c']))  $c['country'] = substr($req['c'], 0, 4);
    if (!empty($req['m']))  $c['method']  = substr($req['m'], 0, 10);
    if (!empty($req['u']))  $c['name']    = substr($req['u'], 0, 40);
    if (!empty($req['sm'])) $c['model']   = substr($req['sm'], 0, 40);   // scanner model
    if (!isset($c['active'])) $c['active']=array();
    if (!in_array($today, $c['active'])) { $c['active'][]=$today; $c['active']=array_slice($c['active'],-90); }
    if ($n) {
        $c['scans'] = intval($c['scans']) + $n;
        if (!isset($c['days'])) $c['days']=array();
        $c['days'][$today] = (isset($c['days'][$today])?intval($c['days'][$today]):0) + $n;
        if (count($c['days'])>120) { ksort($c['days']); $c['days']=array_slice($c['days'],-120,null,true); }
    }
    $d['online'][$client] = $now;
    return true;
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
    $today = today_str(); $now = time();
    $online = 0;
    foreach ($d['online'] as $ts) { if ($now - intval($ts) <= 300) $online++; }

    $week = array(); $month = array(); $bestDay = 0;
    for ($j=6;$j>=0;$j--){ $k=date('Y-m-d',$now-$j*86400); $week[] =array($k,intval(isset($d['days'][$k])?$d['days'][$k]:0)); }
    for ($j=29;$j>=0;$j--){ $k=date('Y-m-d',$now-$j*86400); $month[]=array($k,intval(isset($d['days'][$k])?$d['days'][$k]:0)); }
    foreach ($d['days'] as $v) { if (intval($v) > $bestDay) $bestDay = intval($v); }

    $versions=array(); $countries=array(); $methods=array(); $scores=array();
    $named=array(); $mine=0; $newToday=0;
    foreach ($d['clients'] as $id => $c) {
        if (!empty($c['blocked'])) continue;
        $v =trim(isset($c['version'])?$c['version']:''); if($v!=='')  bump($versions,$v);
        $co=trim(isset($c['country'])?$c['country']:''); if($co!=='') bump($countries,$co);
        $m =trim(isset($c['method'])?$c['method']:'');   if($m!=='')  bump($methods,$m);
        $sc=intval(isset($c['scans'])?$c['scans']:0); $scores[]=$sc;
        $nm=trim(isset($c['name'])?$c['name']:'');
        $named[]=array('name'=>($nm!==''?$nm:'—'),'scans'=>$sc);
        if ($client!=='' && (string)$id===(string)$client) $mine=$sc;
        $fs=intval(isset($c['first'])?$c['first']:0);
        if ($fs && date('Y-m-d',$fs)===$today) $newToday++;
    }
    rsort($scores);
    $rank=1; foreach($scores as $s){ if($s>$mine)$rank++; }
    $top=array_slice($scores,0,10);
    usort($named,function($a,$b){return $b['scans']-$a['scans'];});
    $topNamed=array_slice($named,0,10);

    $todayHours=array_fill(0,24,0); $hourCount=0; $curHK=hour_key();
    foreach ($d['hours'] as $hk=>$hv) {
        if (strpos($hk,$today.'-')===0) { $hh=intval(substr($hk,strlen($today)+1)); if($hh>=0&&$hh<24) $todayHours[$hh]=intval($hv); }
        if ($hk===$curHK) $hourCount=intval($hv);
    }

    // broadcast jo is client par lagoo ho
    $bc = isset($d['broadcast'])?$d['broadcast']:array('msg'=>'','target'=>'all','id'=>0);
    $bmsg=''; $bid=intval(isset($bc['id'])?$bc['id']:0);
    if (!empty($bc['msg'])) {
        $tgt=isset($bc['target'])?$bc['target']:'all';
        $myver=''; if($client!=='' && isset($d['clients'][$client])) $myver=trim($d['clients'][$client]['version']);
        $latest=0; foreach($versions as $vk=>$vv){ if(is_numeric($vk)&&intval($vk)>$latest)$latest=intval($vk); }
        if ($tgt==='all') $bmsg=$bc['msg'];
        else if ($tgt==='old' && $myver!=='' && intval($myver)<$latest) $bmsg=$bc['msg'];
        else if ($tgt!=='all' && $tgt!=='old' && $myver===$tgt) $bmsg=$bc['msg'];
    }

    return array(
        'ok'=>true,'srv'=>'php2','time'=>date('Y-m-d H:i'),'today_key'=>'day_'.$today,
        'total'=>intval($d['total']),
        'today'=>intval(isset($d['days'][$today])?$d['days'][$today]:0),
        'online'=>$online,'users'=>count($d['clients']),'newToday'=>$newToday,
        'week'=>$week,'month'=>$month,'todayHours'=>$todayHours,
        'peak'=>intval(isset($d['peakDay'][$today])?$d['peakDay'][$today]:0),
        'peakAll'=>intval($d['peakAll']),'bestDay'=>$bestDay,'hour'=>$hourCount,
        'imports'=>intval($d['imports']),'prints'=>intval($d['prints']),
        'versions'=>$versions,'countries'=>$countries,'methods'=>$methods,
        'top'=>$top,'topNamed'=>$topNamed,'topscans'=>(count($top)?$top[0]:0),
        'rank'=>$rank,'myscans'=>$mine,
        'broadcast'=>$bmsg,'broadcastId'=>$bid,
        'rconfig'=>isset($d['rconfig'])?$d['rconfig']:array()
    );
}

// helper: 30-din ka cumulative + naye
function build_growth($d, $now) {
    $joinByDay=array();
    foreach ($d['clients'] as $c) { if(!empty($c['blocked']))continue; $fd=date('Y-m-d',intval(isset($c['first'])?$c['first']:0)); $joinByDay[$fd]=(isset($joinByDay[$fd])?$joinByDay[$fd]:0)+1; }
    $newU=array(); $cumU=array();
    for ($i=29;$i>=0;$i--){ $k=date('Y-m-d',$now-$i*86400);
        $newU[]=array($k,isset($joinByDay[$k])?$joinByDay[$k]:0);
        $end=strtotime($k.' 23:59:59'); $cc=0;
        foreach ($d['clients'] as $c){ if(!empty($c['blocked']))continue; if(intval(isset($c['first'])?$c['first']:0)<=$end)$cc++; }
        $cumU[]=array($k,$cc);
    }
    return array($newU,$cumU);
}

// =================================================================
//  CRON:  ?cron=daily&key=CRONKEY   (email report + auto-backup)
//  CLI:   php stats.php cron=daily key=CRONKEY
// =================================================================
if (PHP_SAPI === 'cli' && isset($argv)) { foreach ($argv as $a){ if(strpos($a,'=')!==false){ list($kk,$vv)=explode('=',$a,2); $_GET[$kk]=$vv; } } }
if (isset($_GET['cron'])) {
    if ($_GET['cron']==='daily' && isset($_GET['key']) && hash_equals($CRON_KEY,(string)$_GET['key'])) {
        $d = load_data($DATA_FILE);
        // auto-backup (roz ek, 14 rakho)
        if (!is_dir($BACKUP_DIR)) @mkdir($BACKUP_DIR,0755,true);
        $bf = $BACKUP_DIR.'/stats-'.date('Y-m-d').'.json';
        @file_put_contents($bf, json_encode($d));
        $all=glob($BACKUP_DIR.'/stats-*.json'); if($all && count($all)>14){ sort($all); foreach(array_slice($all,0,count($all)-14) as $old) @unlink($old); }
        $d['lastBackup']=time(); save_data($DATA_FILE,$d);
        // email
        $y=date('Y-m-d',time()-86400); $t=today_str();
        $yc=intval(isset($d['days'][$y])?$d['days'][$y]:0);
        $tc=intval(isset($d['days'][$t])?$d['days'][$t]:0);
        $newY=0; foreach($d['clients'] as $c){ if(date('Y-m-d',intval($c['first']))===$y)$newY++; }
        $msg="ApneScan — Daily Report ($y)\n\n"
            ."Kal ke scans: $yc\nAaj abhi tak: $tc\nTotal (all-time): ".intval($d['total'])."\n"
            ."Total users: ".count($d['clients'])."\nKal naye users: $newY\n"
            ."Crashes stored: ".count($d['crashes'])."\nFeedback stored: ".count($d['feedback'])."\n\n"
            ."Backup: ".basename($bf)."\nDashboard: open ?admin=...\n";
        if ($ADMIN_EMAIL!=='') @mail($ADMIN_EMAIL, "ApneScan daily — $yc scans kal", $msg, "From: ApneScan <no-reply@apnesoft.com>");
        header('Content-Type: text/plain'); echo "cron ok\n".$msg; exit;
    }
    header('Content-Type: text/plain'); echo "cron: bad key"; exit;
}

// ================= ADMIN — login + dashboard =================
if (isset($_GET['admin'])) {
    session_start();
    if (isset($_GET['logout'])) { $_SESSION=array(); @session_destroy(); header('Location: '.strtok($_SERVER['REQUEST_URI'],'?')); exit; }
    // session timeout
    if (!empty($_SESSION['ok']) && isset($_SESSION['seen']) && (time()-$_SESSION['seen'])>$GLOBALS['SESSION_TTL']) { $_SESSION=array(); @session_destroy(); }
    // login
    $login_err='';
    if ($_SERVER['REQUEST_METHOD']==='POST' && isset($_POST['pass'])) {
        $dd=load_data($DATA_FILE); $ip=isset($_SERVER['REMOTE_ADDR'])?$_SERVER['REMOTE_ADDR']:'?';
        $fails=0; foreach((isset($dd['failLog'])?$dd['failLog']:array()) as $f){ if($f['ip']===$ip && time()-$f['t']<600)$fails++; }
        if ($fails>=$MAX_FAILS) { $login_err='Bahut galat koshish. 10 min baad try karo.'; }
        else if (hash_equals($ADMIN_PASS,(string)$_POST['pass'])) { $_SESSION['ok']=true; $_SESSION['seen']=time(); }
        else {
            $login_err='Galat password. Dubara koshish karo.';
            $dd['failLog'][]=array('ip'=>$ip,'t'=>time()); $dd['failLog']=array_slice($dd['failLog'],-30); save_data($DATA_FILE,$dd);
        }
    }
    if (empty($_SESSION['ok'])) {
        header('Content-Type: text/html; charset=utf-8'); ?><!doctype html>
<html lang="hi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ApneScan Admin — Login</title><style>
  *{box-sizing:border-box}
  body{margin:0;font-family:Inter,system-ui,'Segoe UI',Roboto,Arial;background:#0a0f1c;color:#e2e8f0;
    display:flex;align-items:center;justify-content:center;min-height:100vh;position:relative;overflow:hidden}
  body::before{content:"";position:absolute;width:520px;height:520px;border-radius:50%;
    background:radial-gradient(circle,rgba(42,120,214,.45),transparent 62%);top:-140px;left:-120px;filter:blur(20px)}
  body::after{content:"";position:absolute;width:520px;height:520px;border-radius:50%;
    background:radial-gradient(circle,rgba(23,138,138,.42),transparent 62%);bottom:-160px;right:-120px;filter:blur(20px)}
  .box{position:relative;z-index:2;background:rgba(255,255,255,.07);backdrop-filter:blur(16px);
    border:1px solid rgba(255,255,255,.14);color:#fff;width:340px;max-width:90%;padding:34px 30px;
    border-radius:22px;box-shadow:0 24px 70px rgba(0,0,0,.5);text-align:center}
  .box .logo{width:66px;height:66px;margin:0 auto 14px;border-radius:19px;font-size:32px;
    background:linear-gradient(135deg,#2a78d6,#178a8a);display:flex;align-items:center;justify-content:center;
    box-shadow:0 10px 26px rgba(42,120,214,.45)}
  h1{font-size:21px;margin:4px 0 3px;font-weight:800;letter-spacing:-.01em}
  .sub{color:#a9b6ce;font-size:12.5px;margin-bottom:22px}
  input{width:100%;padding:13px 14px;border:1px solid rgba(255,255,255,.18);border-radius:12px;font-size:15px;
    margin-bottom:12px;background:rgba(255,255,255,.08);color:#fff;outline:none;transition:border-color .15s,box-shadow .15s}
  input::placeholder{color:#8ea0bd}
  input:focus{border-color:#3987e5;box-shadow:0 0 0 3px rgba(57,135,229,.28)}
  button{width:100%;padding:13px;border:none;border-radius:12px;background:linear-gradient(135deg,#2a78d6,#1f6fb0);
    color:#fff;font-size:15px;font-weight:800;cursor:pointer;box-shadow:0 8px 22px rgba(42,120,214,.4);transition:filter .15s,transform .1s}
  button:hover{filter:brightness(1.08)} button:active{transform:scale(.98)}
  .err{color:#fca5a5;font-size:12.5px;margin-bottom:10px;background:rgba(227,73,72,.15);padding:8px;border-radius:9px}
  .pv{color:#7f8fab;font-size:11px;margin-top:16px}
</style></head><body>
  <form class="box" method="post" action="">
    <div class="logo">📊</div><h1>ApneScan Admin</h1><div class="sub">Worldwide Analytics Dashboard</div>
    <?php if($login_err) echo '<div class="err">'.htmlspecialchars($login_err).'</div>'; ?>
    <input type="password" name="pass" placeholder="🔒  Password" autofocus required>
    <button type="submit">Login →</button>
    <div class="pv">🔐 Sirf ginti — koi document/patient data nahi</div>
  </form>
</body></html><?php exit;
    }

    // ---- login ho gaya ----
    $_SESSION['seen']=time();
    $t0=microtime(true);
    $d=load_data($DATA_FILE);

    // login log (session me ek baar) — IP ke saath
    if (empty($_SESSION['logged'])) {
        $_SESSION['logged']=true;
        $d['adminLogins'][]=array('t'=>time(),'ip'=>isset($_SERVER['REMOTE_ADDR'])?$_SERVER['REMOTE_ADDR']:'?');
        $d['adminLogins']=array_slice($d['adminLogins'],-20); save_data($DATA_FILE,$d);
    }

    // ---- POST admin actions (user manage / broadcast / config / purge) ----
    if ($_SERVER['REQUEST_METHOD']==='POST' && isset($_POST['act'])) {
        $act=$_POST['act']; $id=isset($_POST['id'])?$_POST['id']:'';
        if ($act==='rename' && isset($d['clients'][$id]))  $d['clients'][$id]['name']=substr(trim($_POST['name']),0,40);
        if ($act==='note'   && isset($d['clients'][$id]))  $d['clients'][$id]['note']=substr(trim($_POST['note']),0,200);
        if ($act==='tag'    && isset($d['clients'][$id]))  $d['clients'][$id]['tags']=substr(trim($_POST['tags']),0,60);
        if ($act==='block'  && isset($d['clients'][$id]))  $d['clients'][$id]['blocked']=!empty($_POST['on'])?1:0;
        if ($act==='broadcast') $d['broadcast']=array('msg'=>substr(trim($_POST['msg']),0,200),'target'=>(isset($_POST['target'])?$_POST['target']:'all'),'id'=>time());
        if ($act==='config') { $cfg=json_decode(isset($_POST['json'])?$_POST['json']:'',true); if(is_array($cfg)) $d['rconfig']=$cfg; }
        if ($act==='purge')  { $days=max(30,intval($_POST['days'])); $cut=date('Y-m-d',time()-$days*86400); foreach($d['days'] as $k=>$v){ if($k<$cut) unset($d['days'][$k]); } }
        if ($act==='clearcrashes') $d['crashes']=array();
        if ($act==='clearfeedback') $d['feedback']=array();
        save_data($DATA_FILE,$d);
        header('Location: '.strtok($_SERVER['REQUEST_URI'],'?').'?admin=1'); exit;
    }

    // ---- EXPORT ----
    if (isset($_GET['export'])) {
        $e=$_GET['export'];
        if ($e==='json'){ header('Content-Type: application/json'); header('Content-Disposition: attachment; filename="apnescan-backup.json"'); echo json_encode($d); exit; }
        if ($e==='days'){ header('Content-Type: text/csv; charset=utf-8'); header('Content-Disposition: attachment; filename="apnescan-daily.csv"'); echo "date,scans\n"; ksort($d['days']); foreach($d['days'] as $dt=>$c) echo $dt.",".intval($c)."\n"; exit; }
        if ($e==='users'){ header('Content-Type: text/csv; charset=utf-8'); header('Content-Disposition: attachment; filename="apnescan-users.csv"'); echo "name,scans,first_seen,last_seen,version,country,method,model,tags\n";
            foreach ($d['clients'] as $c){ echo '"'.str_replace('"','""',trim(isset($c['name'])?$c['name']:'')).'",'.intval(isset($c['scans'])?$c['scans']:0).','.date('Y-m-d H:i',intval(isset($c['first'])?$c['first']:0)).','.date('Y-m-d H:i',intval(isset($c['last'])?$c['last']:0)).','.trim(isset($c['version'])?$c['version']:'').','.trim(isset($c['country'])?$c['country']:'').','.trim(isset($c['method'])?$c['method']:'').',"'.str_replace('"','""',trim(isset($c['model'])?$c['model']:'')).'","'.str_replace('"','""',trim(isset($c['tags'])?$c['tags']:'')).'"'."\n"; }
            exit; }
    }

    // ---- rich compute ----
    $S=compute_stats($d,''); $today=today_str(); $now=time();
    $yest=date('Y-m-d',$now-86400);
    $S['yesterday']=intval(isset($d['days'][$yest])?$d['days'][$yest]:0);
    $wt=0; for($i=0;$i<7;$i++){ $k=date('Y-m-d',$now-$i*86400); $wt+=intval(isset($d['days'][$k])?$d['days'][$k]:0); }
    $mt=0; for($i=0;$i<30;$i++){ $k=date('Y-m-d',$now-$i*86400); $mt+=intval(isset($d['days'][$k])?$d['days'][$k]:0); }
    $S['weekTotal']=$wt; $S['monthTotal']=$mt; $S['dailyAvg']=round($mt/30,1);
    // this month vs last month
    $tm=0; for($i=0;$i<30;$i++){ $k=date('Y-m-d',$now-$i*86400); $tm+=intval(isset($d['days'][$k])?$d['days'][$k]:0); }
    $lm=0; for($i=30;$i<60;$i++){ $k=date('Y-m-d',$now-$i*86400); $lm+=intval(isset($d['days'][$k])?$d['days'][$k]:0); }
    $S['thisMonth']=$tm; $S['lastMonth']=$lm; $S['momPct']=($lm>0)?round(100*($tm-$lm)/$lm):($tm>0?100:0);
    // DAU/WAU/MAU + power/one-time + retention + churn
    $dau=$wau=$mau=0; $power=0; $onetime=0; $churn=array(); $ret1=array('e'=>0,'r'=>0); $ret7=array('e'=>0,'r'=>0); $ret30=array('e'=>0,'r'=>0);
    foreach ($d['clients'] as $id=>$c) {
        if (!empty($c['blocked'])) continue;
        $last=intval(isset($c['last'])?$c['last']:0); $first=intval(isset($c['first'])?$c['first']:0); $sc=intval(isset($c['scans'])?$c['scans']:0);
        if (date('Y-m-d',$last)===$today) $dau++;
        if ($now-$last<=7*86400) $wau++;
        if ($now-$last<=30*86400) $mau++;
        if ($sc>=50) $power++;
        if ($sc<=1) $onetime++;
        if ($now-$last>=14*86400 && $sc>=3) $churn[]=array('name'=>(trim($c['name'])!==''?$c['name']:'—'),'scans'=>$sc,'last'=>$last);
        $act=isset($c['active'])?$c['active']:array();
        // retention: jo user itne din pehle bana, kya wo baad me bhi aaya
        if ($first && $now-$first>=1*86400)  { $ret1['e']++;  foreach($act as $ad){ if($ad>date('Y-m-d',$first)){ $ret1['r']++; break; } } }
        if ($first && $now-$first>=7*86400)  { $ret7['e']++;  $c7=date('Y-m-d',$first+7*86400);  foreach($act as $ad){ if($ad>=$c7 && $ad<=date('Y-m-d',$first+9*86400)){ $ret7['r']++; break; } } }
        if ($first && $now-$first>=30*86400) { $ret30['e']++; $c30=date('Y-m-d',$first+28*86400); foreach($act as $ad){ if($ad>=$c30){ $ret30['r']++; break; } } }
    }
    $S['dau']=$dau; $S['wau']=$wau; $S['mau']=$mau; $S['powerUsers']=$power; $S['oneTime']=$onetime;
    usort($churn,function($a,$b){return $b['scans']-$a['scans'];});
    $S['churn']=array_slice($churn,0,12);
    $S['ret1'] =$ret1['e']?round(100*$ret1['r']/$ret1['e']):0;
    $S['ret7'] =$ret7['e']?round(100*$ret7['r']/$ret7['e']):0;
    $S['ret30']=$ret30['e']?round(100*$ret30['r']/$ret30['e']):0;
    // funnel: users -> jinhone scan kiya -> repeat (>=2 active din)
    $inst=0; $scanned=0; $repeat=0;
    foreach ($d['clients'] as $c){ if(!empty($c['blocked']))continue; $inst++; if(intval($c['scans'])>0)$scanned++; if(count(isset($c['active'])?$c['active']:array())>=2)$repeat++; }
    $S['funnel']=array($inst,$scanned,$repeat);
    // growth
    list($newU,$cumU)=build_growth($d,$now); $S['newUsers30']=$newU; $S['cumUsers30']=$cumU;
    // weekly cohorts (last 8 weeks): kitne bane / abhi tak active
    $coh=array();
    for ($w=7;$w>=0;$w--){ $ws=strtotime('monday this week')-$w*7*86400; $we=$ws+7*86400; $made=0; $still=0;
        foreach ($d['clients'] as $c){ if(!empty($c['blocked']))continue; $f=intval($c['first']); if($f>=$ws && $f<$we){ $made++; if($now-intval($c['last'])<=30*86400)$still++; } }
        $coh[]=array('w'=>date('d M',$ws),'made'=>$made,'still'=>$still,'pct'=>($made?round(100*$still/$made):0));
    }
    $S['cohorts']=$coh;
    // scans by country
    $scByCo=array(); foreach($d['clients'] as $c){ if(!empty($c['blocked']))continue; $co=trim($c['country']); if($co!=='') bump($scByCo,$co,intval($c['scans'])); }
    $S['scansByCountry']=$scByCo;
    // all-time 24h heat
    $hoursAll=array_fill(0,24,0);
    foreach ($d['hours'] as $hk=>$hv){ $p=strrpos($hk,'-'); if($p!==false){ $hh=intval(substr($hk,$p+1)); if($hh>=0&&$hh<24)$hoursAll[$hh]+=intval($hv); } }
    $S['hoursAll']=$hoursAll;
    // user list
    $userList=array(); $onlineNames=array();
    foreach ($d['clients'] as $id=>$c) {
        $nm=trim(isset($c['name'])?$c['name']:''); $last=intval(isset($c['last'])?$c['last']:0);
        $userList[]=array('id'=>(string)$id,'name'=>($nm!==''?$nm:'—'),'scans'=>intval(isset($c['scans'])?$c['scans']:0),
            'first'=>intval(isset($c['first'])?$c['first']:0),'last'=>$last,'version'=>trim(isset($c['version'])?$c['version']:''),
            'country'=>trim(isset($c['country'])?$c['country']:''),'method'=>trim(isset($c['method'])?$c['method']:''),
            'model'=>trim(isset($c['model'])?$c['model']:''),'note'=>trim(isset($c['note'])?$c['note']:''),
            'tags'=>trim(isset($c['tags'])?$c['tags']:''),'blocked'=>!empty($c['blocked'])?1:0,'online'=>(($now-$last)<=300));
        if (($now-$last)<=300 && $nm!=='' && empty($c['blocked'])) $onlineNames[]=$nm;
    }
    usort($userList,function($a,$b){return $b['scans']-$a['scans'];});
    $S['userList']=$userList; $S['onlineNames']=$onlineNames;
    // versions old %
    $vers=$S['versions']; $tv=array_sum($vers); $latest=0; $oldpct=0;
    if ($tv>0){ foreach($vers as $k=>$v){ if(is_numeric($k)&&intval($k)>$latest)$latest=intval($k);} $old=0; foreach($vers as $k=>$v){ if(intval($k)!==$latest)$old+=$v;} $oldpct=round(100*$old/$tv); }
    $S['oldVersionPct']=$oldpct; $S['latestVersion']=$latest;
    $S['unusual']=($S['dailyAvg']>=3 && $S['today']>=3*$S['dailyAvg']);
    // features / scanners / settings
    $S['features']=isset($d['features'])?$d['features']:array();
    $S['scanners']=isset($d['scanners'])?$d['scanners']:array();
    $S['dpis']=isset($d['dpis'])?$d['dpis']:array();
    $S['colors']=isset($d['colors'])?$d['colors']:array();
    $S['sizes']=isset($d['sizes'])?$d['sizes']:array();
    // crashes + feedback
    $S['crashes']=array_reverse(array_slice(isset($d['crashes'])?$d['crashes']:array(),-25));
    $fb=isset($d['feedback'])?$d['feedback']:array();
    $S['feedback']=array_reverse(array_slice($fb,-40));
    $rsum=0;$rn=0; foreach($fb as $f){ if(!empty($f['rating'])){$rsum+=intval($f['rating']);$rn++;} }
    $S['avgRating']=$rn?round($rsum/$rn,1):0; $S['ratingCount']=$rn;
    // broadcast + config current
    $S['bcMsg']=isset($d['broadcast']['msg'])?$d['broadcast']['msg']:''; $S['bcTarget']=isset($d['broadcast']['target'])?$d['broadcast']['target']:'all';
    $S['rconfigStr']=json_encode(isset($d['rconfig'])?$d['rconfig']:array());
    // admin logins + ips
    $al=array(); foreach(array_reverse(array_slice(isset($d['adminLogins'])?$d['adminLogins']:array(),-10)) as $x){ if(is_array($x))$al[]=array(date('Y-m-d H:i',$x['t']),isset($x['ip'])?$x['ip']:'?'); else $al[]=array(date('Y-m-d H:i',$x),'?'); }
    $S['adminLogins']=$al;
    $S['fails']=count(isset($d['failLog'])?$d['failLog']:array());
    $S['daysMap']=$d['days'];
    // health
    $S['fileKB']=file_exists($DATA_FILE)?round(filesize($DATA_FILE)/1024,1):0;
    $S['lastBackup']=intval(isset($d['lastBackup'])?$d['lastBackup']:0);
    $S['respMs']=round((microtime(true)-$t0)*1000);

    $J=json_encode($S);
    ?><!doctype html>
<html lang="hi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ApneScan — Admin Stats</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root{
    --bg:#eef1f6;--card:#ffffff;--card2:#f7f9fc;--fg:#0b1220;--fg2:#334155;--mut:#64748b;--line:#e6eaf0;
    --accent:#2a78d6;--accent2:#1baf7a;
    --sh:0 1px 2px rgba(16,24,40,.06),0 1px 3px rgba(16,24,40,.05);
    --sh2:0 10px 24px rgba(16,24,40,.10),0 2px 6px rgba(16,24,40,.06);
    --hd:linear-gradient(120deg,#12325f 0%,#1f5fb0 55%,#178a8a 100%);
    --radius:16px;
  }
  html[data-th=dark]{
    --bg:#0c0f16;--card:#151a24;--card2:#10141c;--fg:#f1f5f9;--fg2:#cbd5e1;--mut:#8a97ad;--line:#232b3a;
    --accent:#3987e5;--accent2:#199e70;
    --sh:0 1px 2px rgba(0,0,0,.5);
    --sh2:0 12px 28px rgba(0,0,0,.55);
    --hd:linear-gradient(120deg,#0b1a33 0%,#123a6b 55%,#0e4b4b 100%);
  }
  *{box-sizing:border-box}
  body{margin:0;font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:var(--bg);color:var(--fg);-webkit-font-smoothing:antialiased}
  a{color:inherit}
  /* app bar */
  header{background:var(--hd);color:#fff;padding:16px 22px;position:sticky;top:0;z-index:20;box-shadow:0 4px 20px rgba(9,20,45,.25);display:flex;align-items:center;gap:14px;flex-wrap:wrap}
  header .brand{display:flex;align-items:center;gap:11px;font-size:19px;font-weight:800;letter-spacing:-.01em}
  header .brand .logo{width:38px;height:38px;border-radius:11px;background:rgba(255,255,255,.16);display:flex;align-items:center;justify-content:center;font-size:20px;backdrop-filter:blur(4px)}
  header .brand small{display:block;font-size:11px;font-weight:500;opacity:.8;letter-spacing:.02em}
  header .live{display:flex;align-items:center;gap:7px;background:rgba(255,255,255,.14);padding:6px 12px;border-radius:30px;font-size:12px}
  header .live .dot{width:8px;height:8px;border-radius:50%;background:#4ade80;box-shadow:0 0 0 0 rgba(74,222,128,.6);animation:pulse 2s infinite}
  @keyframes pulse{0%{box-shadow:0 0 0 0 rgba(74,222,128,.55)}70%{box-shadow:0 0 0 7px rgba(74,222,128,0)}100%{box-shadow:0 0 0 0 rgba(74,222,128,0)}}
  header .toolbar{margin-left:auto;display:flex;gap:8px;align-items:center}
  .iconbtn{width:38px;height:38px;border-radius:11px;border:none;background:rgba(255,255,255,.14);color:#fff;cursor:pointer;font-size:16px;display:inline-flex;align-items:center;justify-content:center;text-decoration:none;transition:background .15s,transform .1s}
  .iconbtn:hover{background:rgba(255,255,255,.28)} .iconbtn:active{transform:scale(.93)}
  .iconbtn.logout{width:auto;padding:0 14px;gap:6px;font-size:13px;font-weight:700}
  .wrap{max-width:1240px;margin:22px auto;padding:0 18px}
  /* section titles */
  .sec{display:flex;align-items:center;gap:12px;margin:30px 2px 14px;font-size:12px;font-weight:800;letter-spacing:.10em;text-transform:uppercase;color:var(--mut)}
  .sec .em{font-size:15px}
  .sec::after{content:"";flex:1;height:1px;background:linear-gradient(90deg,var(--line),transparent)}
  .sec:first-of-type{margin-top:6px}
  /* KPI */
  .kpis{display:grid;grid-template-columns:repeat(auto-fill,minmax(158px,1fr));gap:13px}
  .kpi{display:flex;align-items:center;gap:13px;background:var(--card);border:1px solid var(--line);border-radius:var(--radius);padding:15px 16px;box-shadow:var(--sh);transition:transform .16s ease,box-shadow .16s ease}
  .kpi:hover{transform:translateY(-3px);box-shadow:var(--sh2)}
  .kpi .ic{width:44px;height:44px;flex:none;border-radius:13px;display:flex;align-items:center;justify-content:center;font-size:21px;background:rgba(42,120,214,.12);color:var(--accent)}
  .kpi .n{font-size:23px;font-weight:800;line-height:1.05;letter-spacing:-.02em}
  .kpi .l{color:var(--mut);font-size:11px;margin-top:4px;font-weight:600;letter-spacing:.01em}
  .kpi.g .ic{background:rgba(27,175,122,.14);color:var(--accent2)} .kpi.g .n{color:var(--accent2)}
  .kpi.r .ic{background:rgba(227,73,72,.13);color:#e34948} .kpi.r .n{color:#e34948}
  .kpi.p .ic{background:rgba(74,58,167,.14);color:#6b5cf0} .kpi.p .n{color:#6b5cf0}
  .kpi.o .ic{background:rgba(235,104,52,.13);color:#eb6834}
  .kpi.y .ic{background:rgba(237,161,0,.15);color:#c98500}
  /* layout */
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
  .grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px}
  @media(max-width:900px){.grid,.grid3{grid-template-columns:1fr}}
  .card{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);padding:16px 17px;margin-bottom:16px;box-shadow:var(--sh)}
  .card h3{margin:0 0 13px;font-size:13.5px;font-weight:700;display:flex;align-items:center;gap:8px;color:var(--fg)}
  .card h3 .em{font-size:16px}
  .card canvas{height:172px!important;width:100%!important}
  table{width:100%;border-collapse:collapse;font-size:13px}
  td,th{padding:7px 8px;border-bottom:1px solid var(--line);text-align:left}
  tbody tr:last-child td{border-bottom:none}
  th{cursor:pointer;color:var(--mut);font-size:11px;text-transform:uppercase;letter-spacing:.04em;font-weight:700}
  .bar{height:9px;background:linear-gradient(90deg,var(--accent),#5fa0ec);border-radius:6px;min-width:6px;box-shadow:inset 0 -1px 2px rgba(0,0,0,.08)}
  /* buttons */
  .rbtn{background:var(--accent);color:#fff;border:none;border-radius:10px;padding:7px 13px;cursor:pointer;font-weight:700;margin-left:6px;text-decoration:none;display:inline-block;font-size:12.5px;box-shadow:var(--sh);transition:filter .15s} .rbtn:hover{filter:brightness(1.07)}
  .rbtn.d{background:var(--card2);color:var(--accent);border:1px solid var(--line)}
  input,select,textarea{padding:9px 11px;border:1px solid var(--line);border-radius:10px;font-size:13px;background:var(--card2);color:var(--fg);font-family:inherit;outline:none;transition:border-color .15s,box-shadow .15s}
  input:focus,select:focus,textarea:focus{border-color:var(--accent);box-shadow:0 0 0 3px rgba(42,120,214,.15)}
  .btn{background:var(--accent);color:#fff;border:none;border-radius:10px;padding:8px 14px;cursor:pointer;font-weight:700;font-size:13px;box-shadow:var(--sh);transition:filter .15s} .btn:hover{filter:brightness(1.07)}
  .btn.gray{background:#64748b}.btn.red{background:#e34948}
  .banner{background:linear-gradient(90deg,#fff7d6,#fef3c7);border:1px solid #fadf8a;color:#8a5a00;border-radius:12px;padding:12px 16px;margin-bottom:14px;font-size:13px;font-weight:600;display:none;box-shadow:var(--sh)}
  html[data-th=dark] .banner{background:linear-gradient(90deg,#3a2f10,#33290c);border-color:#5a4718;color:#fcd34d}
  .btns{margin-bottom:8px;display:flex;gap:2px;flex-wrap:wrap;align-items:center}
  .foot{color:var(--mut);font-size:12px;text-align:center;margin:22px 0 10px}
  .tag{display:inline-block;background:rgba(42,120,214,.12);color:var(--accent);border-radius:20px;padding:2px 9px;font-size:10px;margin-right:3px;font-weight:700}
  .heat{display:grid;grid-template-columns:repeat(24,1fr);gap:3px}
  .heat div{height:30px;border-radius:6px;font-size:8px;color:#fff;text-align:center;line-height:30px;font-weight:600}
  @media print{.toolbar,.btns,.no-print{display:none}.card{break-inside:avoid;box-shadow:none}}
</style></head><body>
<header>
  <div class="brand">
    <div class="logo">📊</div>
    <div>ApneScan <small>Admin Analytics</small></div>
  </div>
  <div class="live no-print"><span class="dot"></span> Live · <span id="tm"></span> · <span id="rt" style="opacity:.75"></span></div>
  <div class="toolbar no-print">
    <button class="iconbtn" onclick="toggleTh()" id="thbtn" title="Theme">🌙</button>
    <button class="iconbtn" onclick="location.reload()" title="Refresh">🔄</button>
    <button class="iconbtn" onclick="window.print()" title="Print">🖨</button>
    <a class="iconbtn logout" href="?logout=1" title="Logout">🔓 Logout</a>
  </div>
</header>
<div class="wrap">
  <div id="banner" class="banner"></div>

  <div class="sec"><span class="em">📊</span> Overview</div>
  <div class="kpis" id="kpis"></div>

  <div class="btns no-print" style="margin-top:14px">
    <span style="color:var(--mut);font-size:12px;font-weight:600;margin-right:4px">📥 Export:</span>
    <a class="rbtn d" href="?admin=1&export=days">Daily CSV</a>
    <a class="rbtn d" href="?admin=1&export=users">Users CSV</a>
    <a class="rbtn d" href="?admin=1&export=json">Backup (JSON)</a>
  </div>

  <div class="sec"><span class="em">📣</span> Broadcast &amp; Control</div>
  <!-- broadcast + remote config -->
  <div class="grid no-print">
    <div class="card"><h3><span class="em">📣</span> Broadcast — sab users ki app me message dikhao</h3>
      <form method="post">
        <input type="hidden" name="act" value="broadcast">
        <input name="msg" maxlength="200" placeholder="Message (jaise: Naya update aa gaya!)" value="<?php echo htmlspecialchars($S['bcMsg']); ?>" style="width:100%;margin-bottom:6px">
        <select name="target" style="margin-right:6px">
          <option value="all"<?php if($S['bcTarget']==='all')echo' selected';?>>Sabhi users</option>
          <option value="old"<?php if($S['bcTarget']==='old')echo' selected';?>>Sirf purane version wale</option>
        </select>
        <button class="btn">Bhejo</button>
        <?php if($S['bcMsg']!==''){ ?><button class="btn gray" name="msg" value="" formnovalidate>Hatao</button><?php } ?>
      </form>
      <div style="color:var(--mut);font-size:11px;margin-top:6px">Abhi: <?php echo $S['bcMsg']!==''?htmlspecialchars($S['bcMsg']).' ('.$S['bcTarget'].')':'— koi message nahi —'; ?></div>
    </div>
    <div class="card"><h3><span class="em">⚙️</span> Remote config — bina update ke app settings badlo</h3>
      <form method="post">
        <input type="hidden" name="act" value="config">
        <textarea name="json" rows="3" style="width:100%" placeholder='{"default_dpi":150}'><?php echo htmlspecialchars($S['rconfigStr']); ?></textarea>
        <div style="margin-top:6px"><button class="btn">Save config (JSON)</button>
        <span style="color:var(--mut);font-size:11px">App ise padhkar apni settings badal sakti hai</span></div>
      </form>
    </div>
  </div>

  <div class="card no-print"><h3><span class="em">🔍</span> Kisi bhi range ka jod</h3>
    <input type="date" id="df"> → <input type="date" id="dt"> <b id="drsum" style="margin-left:10px;color:var(--accent)"></b>
  </div>

  <div class="sec"><span class="em">📈</span> Trends</div>
  <div class="grid">
    <div class="card"><h3><span class="em">📊</span> Last 7 days</h3><canvas id="wk" height="150"></canvas></div>
    <div class="card"><h3><span class="em">📈</span> Last 30 days</h3><canvas id="mo" height="150"></canvas></div>
    <div class="card"><h3><span class="em">🌍</span> Today — 24 hours</h3><canvas id="hr" height="150"></canvas></div>
    <div class="card"><h3><span class="em">🏆</span> Top users (naam ke saath)</h3><canvas id="tp" height="150"></canvas></div>
    <div class="card"><h3><span class="em">🌱</span> New users — 30 days</h3><canvas id="nu" height="150"></canvas></div>
    <div class="card"><h3><span class="em">👥</span> Total users (growth)</h3><canvas id="cu" height="150"></canvas></div>
  </div>

  <div class="sec"><span class="em">🔁</span> Growth &amp; Retention</div>
  <!-- retention + funnel + cohort -->
  <div class="grid3">
    <div class="card"><h3><span class="em">🔁</span> Retention (wapas aaye)</h3><div id="ret"></div></div>
    <div class="card"><h3><span class="em">🫗</span> Funnel</h3><div id="fun"></div></div>
    <div class="card"><h3><span class="em">📅</span> Weekly cohorts (30-din tak tike)</h3><div id="coh"></div></div>
  </div>

  <div class="card"><h3><span class="em">🕐</span> Busiest hours (all-time)</h3><div class="heat" id="heat"></div></div>

  <div class="sec"><span class="em">👤</span> Users</div>
  <div class="card"><h3><span class="em">👤</span> Saare users (<span id="ucount"></span>) <span style="color:var(--mut);font-weight:500;font-size:11px">— header pe click = sort, naam pe click = details/manage</span></h3>
    <input id="usearch" class="no-print" placeholder="🔍 naam / desh / version / tag se dhoondo…" style="width:100%;margin-bottom:8px">
    <div style="overflow:auto;max-height:460px"><table id="utable"></table></div>
  </div>

  <div id="umodal" class="no-print" onclick="if(event.target===this)closeUser()" style="display:none;position:fixed;inset:0;background:rgba(15,23,42,.55);z-index:99;align-items:center;justify-content:center">
    <div style="background:var(--card);color:var(--fg);border-radius:14px;max-width:460px;width:94%;max-height:90vh;overflow:auto;padding:22px;box-shadow:0 20px 60px rgba(0,0,0,.3);position:relative">
      <button onclick="closeUser()" style="position:absolute;top:12px;right:14px;border:none;background:var(--line);border-radius:8px;width:30px;height:30px;cursor:pointer;font-size:16px">✕</button>
      <div id="umbody"></div>
    </div>
  </div>

  <div class="sec"><span class="em">🧰</span> Activity, Usage &amp; Devices</div>
  <div class="grid">
    <div class="card"><h3><span class="em">🟢</span> Abhi online (<span id="oncount"></span>)</h3><div id="onlist"></div></div>
    <div class="card"><h3><span class="em">⏱</span> Recently active</h3><div id="recent"></div></div>
    <div class="card"><h3><span class="em">😴</span> Chhute hue users (churn)</h3><div id="chn"></div></div>
    <div class="card"><h3><span class="em">🧰</span> Feature usage</h3><div id="ft"></div></div>
    <div class="card"><h3><span class="em">🖨</span> Scanner models</h3><div id="sm"></div></div>
    <div class="card"><h3><span class="em">🎚</span> Scan settings (DPI / colour / size)</h3><div id="ss"></div></div>
    <div class="card"><h3><span class="em">🗺</span> Countries (users)</h3><div id="co"></div></div>
    <div class="card"><h3><span class="em">🗺</span> Scans by country</h3><div id="cos"></div></div>
    <div class="card"><h3><span class="em">🔢</span> App versions</h3><div id="ve"></div></div>
    <div class="card"><h3><span class="em">🖨</span> Scan methods</h3><div id="me"></div></div>
  </div>

  <div class="sec"><span class="em">💬</span> Quality, Feedback &amp; System</div>
  <div class="grid">
    <div class="card"><h3><span class="em">💬</span> Feedback (⭐ <span id="arate"></span>)</h3><div id="fb"></div></div>
    <div class="card"><h3><span class="em">💥</span> Crash reports</h3><div id="cr"></div></div>
    <div class="card"><h3><span class="em">🏅</span> Records</h3><div id="rc"></div></div>
    <div class="card"><h3><span class="em">🔒</span> Admin logins (IP)</h3><div id="al"></div></div>
  </div>

  <!-- maintenance -->
  <div class="card no-print"><h3><span class="em">🛠</span> Maintenance</h3>
    <form method="post" style="display:inline">
      <input type="hidden" name="act" value="purge">
      <button class="btn gray" onclick="return confirm('Purane din ka data hata dein?')">Purane data hatao (rakho last</button>
      <select name="days"><option value="90">90</option><option value="180">180</option><option value="365">365</option></select> <span style="color:var(--mut)">din)</span>
    </form>
    <form method="post" style="display:inline;margin-left:8px"><input type="hidden" name="act" value="clearcrashes"><button class="btn gray">Crashes clear</button></form>
    <form method="post" style="display:inline"><input type="hidden" name="act" value="clearfeedback"><button class="btn gray">Feedback clear</button></form>
    <div style="color:var(--mut);font-size:11px;margin-top:8px">Data file: <b><?php echo $S['fileKB']; ?> KB</b> · Backup: <b><?php echo $S['lastBackup']?date('d M H:i',$S['lastBackup']):'—'; ?></b> · Fail-logins: <b><?php echo $S['fails']; ?></b></div>
  </div>

  <div class="foot">Server: PHP <?php echo htmlspecialchars($S['srv']); ?> · <?php echo htmlspecialchars($S['time']); ?> · ApneSoftware.com</div>
</div>
<script>
var D=<?php echo $J; ?>;
document.getElementById('tm').textContent=D.time;
document.getElementById('rt').textContent='server '+D.respMs+'ms';
function flag(cc){ if(!cc||cc.length!==2) return '🏳'; return String.fromCodePoint.apply(null,[].map.call(cc.toUpperCase(),function(ch){return 127397+ch.charCodeAt(0);})); }
function fmt(n){ return (n||0).toLocaleString(); }
function esc(s){ return (s||'').replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];}); }
function ago(ts){ if(!ts)return '—'; var s=Math.floor(Date.now()/1000)-ts; if(s<60)return s+'s'; if(s<3600)return Math.floor(s/60)+'m'; if(s<86400)return Math.floor(s/3600)+'h'; return Math.floor(s/86400)+'d'; }

// theme
function toggleTh(){ var dk=document.documentElement.getAttribute('data-th')==='dark'; try{localStorage.setItem('anth',dk?'light':'dark');}catch(e){} location.reload(); }
(function(){ try{ if(localStorage.getItem('anth')==='dark'){ document.documentElement.setAttribute('data-th','dark'); document.getElementById('thbtn').textContent='☀️'; } }catch(e){} })();

// palette (validated categorical) + Chart.js theme
var DKT=document.documentElement.getAttribute('data-th')==='dark';
var PAL=DKT?{blue:'#3987e5',orange:'#d95926',aqua:'#199e70',yellow:'#c98500',magenta:'#d55181',green:'#008300',violet:'#9085e9',red:'#e66767'}
           :{blue:'#2a78d6',orange:'#eb6834',aqua:'#1baf7a',yellow:'#eda100',magenta:'#e87ba4',green:'#008300',violet:'#4a3aa7',red:'#e34948'};
if(window.Chart){
  Chart.defaults.font.family="Inter,system-ui,'Segoe UI',Roboto,Arial";
  Chart.defaults.font.size=11;
  Chart.defaults.color=DKT?'#8a97ad':'#64748b';
  Chart.defaults.borderColor=DKT?'rgba(255,255,255,.06)':'rgba(15,23,42,.06)';
  Chart.defaults.plugins.tooltip.backgroundColor=DKT?'#0b0f16':'#0b1220';
  Chart.defaults.plugins.tooltip.padding=10;
  Chart.defaults.plugins.tooltip.cornerRadius=9;
  Chart.defaults.plugins.tooltip.titleFont={weight:'700'};
  Chart.defaults.maintainAspectRatio=false;
  Chart.defaults.elements.bar.borderRadius=6;
  Chart.defaults.elements.bar.borderSkipped=false;
}
function grad(ctx,hex){try{var c=ctx.chart.ctx,g=c.createLinearGradient(0,0,0,ctx.chart.height||160);g.addColorStop(0,hex+'55');g.addColorStop(1,hex+'05');return g;}catch(e){return hex+'22';}}

// KPIs (icon + number + label)
function kpi(ic,n,l,cls){return '<div class="kpi '+(cls||'')+'"><div class="ic">'+ic+'</div><div><div class="n">'+n+'</div><div class="l">'+l+'</div></div></div>';}
document.getElementById('kpis').innerHTML=
  kpi('📄',fmt(D.total),'Total scans')+kpi('📅',fmt(D.today),'Today')+kpi('🕐',fmt(D.yesterday),'Yesterday')+
  kpi('🗓',fmt(D.weekTotal),'This week')+kpi('📆',fmt(D.monthTotal),'This month')+kpi('📊',D.dailyAvg,'Daily avg')+
  kpi(D.momPct>=0?'📈':'📉',(D.momPct>=0?'+':'')+D.momPct+'%','vs last month',D.momPct>=0?'g':'r')+
  kpi('🟢',fmt(D.online),'Online now','g')+kpi('👥',fmt(D.users),'Total users')+kpi('✨',fmt(D.newToday),'New today','o')+
  kpi('☀️',fmt(D.dau),'Active today')+kpi('📅',fmt(D.wau),'Active 7d')+kpi('🗓',fmt(D.mau),'Active 30d')+
  kpi('🔥',fmt(D.powerUsers),'Power users','p')+kpi('1️⃣',fmt(D.oneTime),'One-time','r')+
  kpi('📥',fmt(D.imports),'Imports')+kpi('🖨',fmt(D.prints),'Prints')+kpi('🌿',fmt(D.total),'Paper saved','g');

// banner
var bmsg='';
if(D.unusual) bmsg='🚀 Aaj normal se bahut zyada activity ho rahi hai!';
[100000,50000,10000,5000,1000,500,100].some(function(m){ if(D.total>=m && D.total-D.today<m){ bmsg='🎉 Milestone! Duniya bhar me '+fmt(m)+' scans paar!'; return true;} return false;});
if(D.oldVersionPct>=30) bmsg=(bmsg?bmsg+'  ':'')+'🔔 '+D.oldVersionPct+'% users purane version par.';
if(D.crashes&&D.crashes.length) bmsg=(bmsg?bmsg+'  ':'')+'💥 '+D.crashes.length+' crash report aaye hain — neeche dekho.';
if(bmsg){ var b=document.getElementById('banner'); b.textContent=bmsg; b.style.display='block'; }

// bars helper
function bars(id,rows,useFlag,label){var mx=rows[0]?rows[0][1]:1;if(!rows.length){document.getElementById(id).innerHTML='<div style="color:var(--mut);font-size:12px">— abhi koi data nahi —</div>';return;}var h='<table>';rows.forEach(function(r){
  var lab=useFlag?(flag(r[0])+' '+r[0]):(label?label(r[0]):esc(r[0]));
  h+='<tr><td style="width:120px">'+lab+'</td><td><div class="bar" style="width:'+Math.max(4,100*r[1]/(mx||1))+'%"></div></td><td style="width:44px;text-align:right">'+r[1]+'</td></tr>';});
  document.getElementById(id).innerHTML=h+'</table>';}
function obj2rows(o){return Object.keys(o||{}).map(function(k){return [k,o[k]];}).sort(function(a,b){return b[1]-a[1];}).slice(0,12);}

// charts (agar Chart.js load na ho to baaki dashboard fir bhi chale)
function mkChart(el,cfg){ try{ if(window.Chart) return new Chart(el,cfg); }catch(e){}
  try{ var c=document.getElementById(el.id||el); if(c&&c.parentNode){ var d=document.createElement('div'); d.style.cssText='color:var(--mut);font-size:12px;text-align:center;padding:20px'; d.textContent='📊 chart load nahi hua (internet?)'; c.parentNode.replaceChild(d,c); } }catch(e){} return null; }
if(window.Chart){
  var bc=Chart.defaults.borderColor;
  var CO={plugins:{legend:{display:false}},scales:{x:{grid:{display:false}},y:{beginAtZero:true,grid:{color:bc},ticks:{precision:0}}}};
  var COy={indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{beginAtZero:true,grid:{color:bc}},y:{grid:{display:false}}}};
  mkChart(wk,{type:'bar',data:{labels:D.week.map(function(x){return x[0].slice(5);}),datasets:[{data:D.week.map(function(x){return x[1];}),backgroundColor:PAL.blue}]},options:CO});
  mkChart(mo,{type:'line',data:{labels:D.month.map(function(x){return x[0].slice(5);}),datasets:[{data:D.month.map(function(x){return x[1];}),borderColor:PAL.blue,backgroundColor:function(c){return grad(c,PAL.blue);},borderWidth:2,fill:true,tension:.35,pointRadius:0}]},options:CO});
  mkChart(hr,{type:'bar',data:{labels:D.todayHours.map(function(_,i){return i;}),datasets:[{data:D.todayHours,backgroundColor:PAL.aqua}]},options:CO});
  var TN=(D.topNamed&&D.topNamed.length)?D.topNamed:(D.top||[]).map(function(s,i){return {name:'User '+(i+1),scans:s};});
  mkChart(tp,{type:'bar',data:{labels:TN.map(function(t,i){return (t.name&&t.name!=='—')?t.name:('User '+(i+1));}),datasets:[{data:TN.map(function(t){return t.scans;}),backgroundColor:PAL.violet}]},options:COy});
  mkChart(nu,{type:'bar',data:{labels:D.newUsers30.map(function(x){return x[0].slice(5);}),datasets:[{data:D.newUsers30.map(function(x){return x[1];}),backgroundColor:PAL.aqua}]},options:CO});
  mkChart(cu,{type:'line',data:{labels:D.cumUsers30.map(function(x){return x[0].slice(5);}),datasets:[{data:D.cumUsers30.map(function(x){return x[1];}),borderColor:PAL.violet,backgroundColor:function(c){return grad(c,PAL.violet);},borderWidth:2,fill:true,tension:.35,pointRadius:0}]},options:CO});
}

// retention / funnel / cohorts
function retRow(l,p){ return '<tr><td>'+l+'</td><td style="width:55%"><div class="bar" style="width:'+Math.max(3,p)+'%;background:'+PAL.aqua+'"></div></td><td style="text-align:right"><b>'+p+'%</b></td></tr>'; }
document.getElementById('ret').innerHTML='<table>'+retRow('Next day (D+1)',D.ret1)+retRow('Week 1 (D+7)',D.ret7)+retRow('Month 1 (D+30)',D.ret30)+'</table>';
var fn=D.funnel||[0,0,0]; function funRow(l,v,base){ var p=base?Math.round(100*v/base):0; return '<tr><td>'+l+'</td><td style="width:50%"><div class="bar" style="width:'+Math.max(3,p)+'%;background:'+PAL.blue+'"></div></td><td style="text-align:right"><b>'+v+'</b> ('+p+'%)</td></tr>'; }
document.getElementById('fun').innerHTML='<table>'+funRow('Installed',fn[0],fn[0])+funRow('Scanned ≥1',fn[1],fn[0])+funRow('Repeat (2+ din)',fn[2],fn[0])+'</table>';
document.getElementById('coh').innerHTML='<table><tr><th>Week</th><th>Naye</th><th>Tike</th></tr>'+(D.cohorts||[]).map(function(c){return '<tr><td>'+c.w+'</td><td>'+c.made+'</td><td>'+c.still+' <span style="color:var(--mut)">('+c.pct+'%)</span></td></tr>';}).join('')+'</table>';

// heat (all-time hours)
var hmax=Math.max.apply(null,D.hoursAll.concat([1]));
document.getElementById('heat').innerHTML=D.hoursAll.map(function(v,i){var a=v/hmax;var bg='rgba(42,120,214,'+(0.10+a*0.90)+')';return '<div title="'+i+':00 — '+v+' scans" style="background:'+bg+';color:'+(a>0.45?'#fff':'var(--mut)')+'">'+i+'</div>';}).join('');

// breakdowns
bars('co',obj2rows(D.countries),true);
bars('cos',obj2rows(D.scansByCountry),true);
bars('ve',obj2rows(D.versions),false,function(k){return 'v'+k;});
bars('me',Object.keys(D.methods||{}).map(function(k){var n={escl:'Network',wia:'USB',twain:'TWAIN',naps2:'NAPS2'}[k]||k;return [n,D.methods[k]];}).sort(function(a,b){return b[1]-a[1];}),false);
bars('ft',obj2rows(D.features),false,function(k){var n={ocr:'OCR',compress:'Compress',share:'Share',whatsapp:'WhatsApp',email:'Email',print:'Print',import:'Import',merge:'Merge',split:'Split',rotate:'Rotate',sign:'Sign',protect:'Password'}[k];return n||k;});
bars('sm',obj2rows(D.scanners),false);
// scan settings combined
(function(){var box=document.getElementById('ss');var h='';
  function seg(t,o){var r=obj2rows(o);if(!r.length)return '';var mx=r[0][1];return '<div style="font-weight:600;font-size:12px;margin:6px 0 3px">'+t+'</div><table>'+r.map(function(x){return '<tr><td style="width:90px">'+esc(x[0])+'</td><td><div class="bar" style="width:'+Math.max(4,100*x[1]/mx)+'%"></div></td><td style="width:40px;text-align:right">'+x[1]+'</td></tr>';}).join('')+'</table>';}
  h+=seg('DPI',D.dpis)+seg('Colour',D.colors)+seg('Page size',D.sizes);
  box.innerHTML=h||'<div style="color:var(--mut);font-size:12px">— abhi koi data nahi (naye app version se aayega) —</div>';})();

// online + recent + churn
document.getElementById('oncount').textContent=(D.onlineNames||[]).length;
document.getElementById('onlist').innerHTML=(D.onlineNames&&D.onlineNames.length)?D.onlineNames.map(function(n){return '<div>🟢 '+esc(n)+'</div>';}).join(''):'<div style="color:var(--mut)">Abhi koi online nahi</div>';
var rec=(D.userList||[]).slice().sort(function(a,b){return b.last-a.last;}).slice(0,10);
document.getElementById('recent').innerHTML='<table>'+rec.map(function(u){return '<tr><td>'+(u.online?'🟢 ':'⚪ ')+esc(u.name)+'</td><td style="text-align:right;color:var(--mut)">'+ago(u.last)+' ago</td></tr>';}).join('')+'</table>';
document.getElementById('chn').innerHTML=(D.churn&&D.churn.length)?'<table>'+D.churn.map(function(u){return '<tr><td>😴 '+esc(u.name)+'</td><td style="text-align:right"><b>'+u.scans+'</b> · '+ago(u.last)+' pehle</td></tr>';}).join('')+'</table>':'<div style="color:var(--mut);font-size:12px">— sab active hain 🎉 —</div>';

// feedback + crashes
document.getElementById('arate').textContent=(D.avgRating||0)+' ('+(D.ratingCount||0)+')';
document.getElementById('fb').innerHTML=(D.feedback&&D.feedback.length)?D.feedback.map(function(f){return '<div style="border-bottom:1px solid var(--line);padding:5px 0"><b>'+('★'.repeat(f.rating||0)||'—')+'</b> '+esc(f.name||'—')+' <span style="color:var(--mut);font-size:11px">'+(f.v?'v'+f.v:'')+'</span><div>'+esc(f.msg||'')+'</div></div>';}).join(''):'<div style="color:var(--mut);font-size:12px">— abhi koi feedback nahi —</div>';
document.getElementById('cr').innerHTML=(D.crashes&&D.crashes.length)?'<table>'+D.crashes.map(function(c){return '<tr><td>💥 '+esc((c.err||'').slice(0,60))+'</td><td style="text-align:right;color:var(--mut);white-space:nowrap">v'+(c.v||'?')+' · '+ago(c.t)+'</td></tr>';}).join('')+'</table>':'<div style="color:var(--mut);font-size:12px">— koi crash nahi 🎉 —</div>';

// records + admin logins
document.getElementById('rc').innerHTML='<table>'+
 '<tr><td>🏔 Peak online (all-time)</td><td style="text-align:right"><b>'+D.peakAll+'</b></td></tr>'+
 '<tr><td>📅 Best single day</td><td style="text-align:right"><b>'+D.bestDay+'</b></td></tr>'+
 '<tr><td>🕐 This hour</td><td style="text-align:right"><b>'+D.hour+'</b></td></tr>'+
 '<tr><td>🆕 Latest version</td><td style="text-align:right"><b>v'+D.latestVersion+'</b></td></tr></table>';
document.getElementById('al').innerHTML='<table>'+(D.adminLogins||[]).map(function(t){return '<tr><td>🔒 '+t[0]+'</td><td style="text-align:right;color:var(--mut)">'+t[1]+'</td></tr>';}).join('')+'</table>';

// ---- user table ----
var sortKey='scans', sortDir=-1;
function renderUsers(){
  var q=(document.getElementById('usearch').value||'').toLowerCase();
  var rows=(D.userList||[]).filter(function(u){ return (u.name+' '+u.country+' '+u.version+' '+u.method+' '+u.tags+' '+u.model).toLowerCase().indexOf(q)>=0; });
  rows.sort(function(a,b){ var x=a[sortKey],y=b[sortKey]; if(typeof x==='string'){return sortDir*x.localeCompare(y);} return sortDir*((x||0)-(y||0)); });
  document.getElementById('ucount').textContent=rows.length;
  var h='<tr>'+[['name','Name'],['scans','Scans'],['last','Last seen'],['first','Joined'],['version','Ver'],['country','Desh'],['method','Method']].map(function(c){return '<th data-k="'+c[0]+'">'+c[1]+(sortKey===c[0]?(sortDir<0?' ▼':' ▲'):'')+'</th>';}).join('')+'</tr>';
  rows.forEach(function(u,i){ h+='<tr data-i="'+i+'" style="cursor:pointer'+(u.blocked?';opacity:.45':'')+'"><td>'+(u.blocked?'🚫 ':(u.online?'🟢 ':''))+esc(u.name)+(u.tags?' '+u.tags.split(',').map(function(t){return '<span class="tag">'+esc(t.trim())+'</span>';}).join(''):'')+'</td><td><b>'+u.scans+'</b></td><td>'+ago(u.last)+' ago</td><td>'+(u.first?new Date(u.first*1000).toISOString().slice(0,10):'—')+'</td><td>'+(u.version||'—')+'</td><td>'+(u.country?flag(u.country)+' '+u.country:'—')+'</td><td>'+({escl:'Network',wia:'USB',twain:'TWAIN',naps2:'NAPS2'}[u.method]||u.method||'—')+'</td></tr>'; });
  var t=document.getElementById('utable'); t.innerHTML=h;
  [].forEach.call(t.querySelectorAll('th'),function(th){ th.onclick=function(){ var k=th.getAttribute('data-k'); if(sortKey===k)sortDir*=-1; else {sortKey=k;sortDir=(k==='name'||k==='country'||k==='version'||k==='method')?1:-1;} renderUsers(); }; });
  [].forEach.call(t.querySelectorAll('tr[data-i]'),function(tr){ tr.onclick=function(){ showUser(rows[+tr.getAttribute('data-i')]); }; });
}
function showUser(u){
  if(!u) return;
  var mth={escl:'Network (eSCL/WiFi)',wia:'USB (WIA)',twain:'TWAIN',naps2:'NAPS2'}[u.method]||u.method||'—';
  var share=(D.total>0)?((u.scans*100/D.total).toFixed(u.scans*100/D.total<1?2:1)+'%'):'—';
  function row(l,v){ return '<tr><td style="color:var(--mut);padding:5px 14px 5px 0;white-space:nowrap">'+l+'</td><td style="font-weight:600">'+v+'</td></tr>'; }
  var html='<div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">'+
      '<div style="font-size:26px">'+(u.blocked?'🚫':(u.online?'🟢':'👤'))+'</div>'+
      '<div><div style="font-size:20px;font-weight:700">'+esc(u.name!=='—'?u.name:'(bina naam)')+'</div>'+
      '<div style="color:var(--mut);font-size:13px">'+(u.online?'Abhi online':'Offline · '+ago(u.last)+' pehle')+'</div></div></div>'+
    '<table style="width:100%;font-size:14px">'+
      row('Total scans','<span style="font-size:18px">'+fmt(u.scans)+'</span>')+
      row('World me hissa',share)+
      row('Country',u.country?flag(u.country)+' '+u.country:'—')+
      row('App version',u.version?'v'+u.version:'—')+
      row('Scan method',mth)+
      row('Scanner model',esc(u.model||'—'))+
      row('Joined',u.first?new Date(u.first*1000).toLocaleString():'—')+
      row('Aakhri activity',u.last?new Date(u.last*1000).toLocaleString():'—')+
      (u.note?row('Note','<i>'+esc(u.note)+'</i>'):'')+
    '</table>'+
    '<div style="border-top:1px solid var(--line);margin-top:12px;padding-top:12px">'+
      '<div style="font-weight:700;font-size:13px;margin-bottom:8px">✏️ Manage</div>'+
      '<form method="post" style="margin-bottom:6px"><input type="hidden" name="act" value="rename"><input type="hidden" name="id" value="'+esc(u.id)+'"><input name="name" value="'+esc(u.name!=='—'?u.name:'')+'" placeholder="Naam badlo" style="width:60%"> <button class="btn">Rename</button></form>'+
      '<form method="post" style="margin-bottom:6px"><input type="hidden" name="act" value="tag"><input type="hidden" name="id" value="'+esc(u.id)+'"><input name="tags" value="'+esc(u.tags)+'" placeholder="Tags (comma se): VIP,Hospital" style="width:60%"> <button class="btn">Tag</button></form>'+
      '<form method="post" style="margin-bottom:6px"><input type="hidden" name="act" value="note"><input type="hidden" name="id" value="'+esc(u.id)+'"><input name="note" value="'+esc(u.note)+'" placeholder="Note likho" style="width:60%"> <button class="btn">Note</button></form>'+
      '<form method="post"><input type="hidden" name="act" value="block"><input type="hidden" name="id" value="'+esc(u.id)+'"><input type="hidden" name="on" value="'+(u.blocked?'':'1')+'"><button class="btn '+(u.blocked?'gray':'red')+'">'+(u.blocked?'✅ Unblock':'🚫 Block (stats se hatao)')+'</button></form>'+
    '</div>';
  document.getElementById('umbody').innerHTML=html;
  document.getElementById('umodal').style.display='flex';
}
function closeUser(){ document.getElementById('umodal').style.display='none'; }
document.getElementById('usearch').addEventListener('input',renderUsers);
renderUsers();

// date-range
function drCalc(){ var a=document.getElementById('df').value,b=document.getElementById('dt').value; if(!a||!b){document.getElementById('drsum').textContent='';return;} var s=0; for(var k in D.daysMap){ if(k>=a&&k<=b) s+=parseInt(D.daysMap[k])||0; } document.getElementById('drsum').textContent=fmt(s)+' scans'; }
(function(){var t=new Date(),f=new Date(t.getTime()-6*86400000); document.getElementById('dt').value=t.toISOString().slice(0,10); document.getElementById('df').value=f.toISOString().slice(0,10);})();
document.getElementById('df').addEventListener('change',drCalc);
document.getElementById('dt').addEventListener('change',drCalc); drCalc();

// auto-refresh 30s (agar koi modal/form khula nahi)
setInterval(function(){ if(document.getElementById('umodal').style.display==='flex')return; if(document.activeElement&&/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName))return; location.reload(); },30000);
</script>
</body></html><?php exit;
}

// ================= API (App se) =================
$action = isset($_REQUEST['action']) ? $_REQUEST['action'] : 'stats';
$client = isset($_REQUEST['client']) ? $_REQUEST['client'] : '';
$today  = today_str(); $now = time();
$d = load_data($DATA_FILE);

if ($action === 'scan') {
    if ($SECRET !== '' && (!isset($_REQUEST['secret']) || $_REQUEST['secret'] !== $SECRET)) {
        header('Content-Type: application/json'); echo json_encode(array('ok'=>false,'error'=>'bad secret')); exit;
    }
    $n = max(0, min(100, intval(isset($_REQUEST['n'])?$_REQUEST['n']:1)));
    $okc = touch_client($d, $client, $_REQUEST, $n, $now, $today);   // blocked -> false
    if ($okc) {
        $d['total'] = intval($d['total']) + $n;
        $d['days'][$today] = (isset($d['days'][$today])?intval($d['days'][$today]):0) + $n;
        $hk = hour_key(); $d['hours'][$hk] = (isset($d['hours'][$hk])?intval($d['hours'][$hk]):0) + $n;
        // scan settings breakdown
        if (isset($_REQUEST['dpi'])) bump($d['dpis'],  substr($_REQUEST['dpi'],0,6), $n?1:1);
        if (isset($_REQUEST['col'])) bump($d['colors'],substr($_REQUEST['col'],0,10),1);
        if (isset($_REQUEST['sz']))  bump($d['sizes'], substr($_REQUEST['sz'],0,10),1);
        if (!empty($_REQUEST['sm']))  bump($d['scanners'], substr($_REQUEST['sm'],0,40),1);
        update_peak($d, $now, $today);
    }
} else if ($action === 'ping') {
    touch_client($d, $client, $_REQUEST, 0, $now, $today);
    update_peak($d, $now, $today);
} else if ($action === 'event') {
    touch_client($d, $client, $_REQUEST, 0, $now, $today);
    if (!empty($_REQUEST['feat'])) bump($d['features'], substr($_REQUEST['feat'],0,20), 1);
} else if ($action === 'crash') {
    $d['crashes'][] = array('t'=>$now,'v'=>substr(isset($_REQUEST['v'])?$_REQUEST['v']:'',0,10),
        'err'=>substr(isset($_REQUEST['err'])?$_REQUEST['err']:'',0,200),'client'=>substr($client,0,40));
    $d['crashes'] = array_slice($d['crashes'], -100);
} else if ($action === 'feedback') {
    $d['feedback'][] = array('t'=>$now,'name'=>substr(isset($_REQUEST['u'])?$_REQUEST['u']:'',0,40),
        'v'=>substr(isset($_REQUEST['v'])?$_REQUEST['v']:'',0,10),'rating'=>max(0,min(5,intval(isset($_REQUEST['rating'])?$_REQUEST['rating']:0))),
        'msg'=>substr(isset($_REQUEST['msg'])?$_REQUEST['msg']:'',0,400));
    $d['feedback'] = array_slice($d['feedback'], -300);
    touch_client($d, $client, $_REQUEST, 0, $now, $today);
}

// import/print kisi bhi action ke saath
$imp = max(0, min(500, intval(isset($_REQUEST['imp'])?$_REQUEST['imp']:0)));
$prt = max(0, min(500, intval(isset($_REQUEST['prt'])?$_REQUEST['prt']:0)));
if ($imp) $d['imports'] = intval($d['imports']) + $imp;
if ($prt) $d['prints']  = intval($d['prints'])  + $prt;

foreach ($d['online'] as $id => $ts) { if ($now - intval($ts) > 86400) unset($d['online'][$id]); }
save_data($DATA_FILE, $d);

header('Content-Type: application/json');
echo json_encode(compute_stats($d, $client));
