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
$ADMIN_EMAIL = '';                   // <-- daily/weekly report yahan aayega (khaali = band)
$CRON_KEY    = 'apnecron';           // <-- ?cron=daily&key=... ka key
$SECRET      = '';                   // optional: scan protect karne ko
$SESSION_TTL = 1800;                 // admin auto-logout (sec) — 30 min
$MAX_FAILS   = 6;                    // itni galat koshish -> thodi der lock
// Telegram phone-notification (optional): @BotFather se token banao, apni
// chat id daalo. Dono bhare to milestone/crash/daily par phone par message.
$TELEGRAM_TOKEN = '';                // <-- Telegram bot token (khaali = band)
$TELEGRAM_CHAT  = '';                // <-- aapki chat id
$SITE_NAME      = 'ApneScan';        // public page / widget par dikhega

header('Access-Control-Allow-Origin: *');

// ---------------- storage helpers ----------------
function default_data() {
    return array(
        'total'=>0,'days'=>array(),'hours'=>array(),'imports'=>0,'prints'=>0,
        'peakAll'=>0,'peakDay'=>array(),'clients'=>array(),'online'=>array(),
        'features'=>array(),'scanners'=>array(),'dpis'=>array(),
        'colors'=>array(),'sizes'=>array(),'crashes'=>array(),'feedback'=>array(),
        'broadcast'=>array('msg'=>'','target'=>'all','id'=>0),'rconfig'=>array(),
        'adminLogins'=>array(),'failLog'=>array(),'lastBackup'=>0,
        'scanEvents'=>0,'recentScans'=>array(),'metrics'=>array(),'featUsers'=>array()
    );
}
// ek file se saaf JSON padho (khaali/tuti -> null)
function _read_json($f) {
    if (!is_file($f)) return null;
    $s = @file_get_contents($f);
    if ($s === false) return null;
    $s = trim($s);
    if ($s === '') return null;
    $d = json_decode($s, true);
    return is_array($d) ? $d : null;
}
// kisi bhi data ka "kitna bhara hai" score (jitna zyada, utna asli)
function _data_score($d) {
    if (!is_array($d)) return -1;
    return intval(isset($d['total'])?$d['total']:0)
         + count(isset($d['clients'])&&is_array($d['clients'])?$d['clients']:array())
         + count(isset($d['days'])&&is_array($d['days'])?$d['days']:array());
}
/**
 * BULLETPROOF LOAD: main file, uska shadow (.bak) aur haal ke dated backups —
 * in sab me se jisme SABSE ZYADA data ho wahi lo. Isse agar main file kabhi
 * khaali/tuti/0 ho jaaye to bhi purana data apne aap wapas aa jaata hai.
 */
function load_data($file) {
    $cands = array();
    $m = _read_json($file);          if ($m !== null) $cands[] = $m;
    $b = _read_json($file . '.bak'); if ($b !== null) $cands[] = $b;
    $g = @glob(dirname($file) . '/backups/stats-*.json');
    if ($g) { rsort($g); $n = 0; foreach ($g as $bf) { $c = _read_json($bf); if ($c !== null) { $cands[] = $c; if (++$n >= 3) break; } } }
    if (!$cands) return default_data();
    $best = $cands[0]; $bs = _data_score($best);
    foreach ($cands as $c) { $s = _data_score($c); if ($s > $bs) { $bs = $s; $best = $c; } }
    return array_merge(default_data(), $best);
}
/**
 * BULLETPROOF SAVE:
 *  - counters (total/imports/prints) kabhi kam nahi hote (monotonic)
 *  - non-force save purani days/clients ko wipe nahi karta
 *  - agar naya data KHAALI hai par purana bhara tha -> likhta hi nahi (0 se bachao)
 *  - atomic write (temp file -> rename) taaki aadha-likha/tuta file kabhi na bane
 *  - har safal write ke baad ek shadow copy (.bak)
 * $force=true sirf admin ke jaan-boojhkar kiye actions ke liye (purge waghairah).
 */
function save_data($file, $d, $force = false) {
    $prev = _read_json($file);
    if ($prev === null) $prev = _read_json($file . '.bak');
    if (is_array($prev)) {
        // counters kabhi peeche nahi jaate
        foreach (array('total','imports','prints','peakAll') as $k) {
            if (intval(isset($d[$k])?$d[$k]:0) < intval(isset($prev[$k])?$prev[$k]:0)) $d[$k] = intval($prev[$k]);
        }
        if (!$force) {
            // din ka itihaas wipe se bachao
            if (isset($prev['days']) && is_array($prev['days'])) {
                if (!isset($d['days']) || !is_array($d['days'])) $d['days'] = array();
                foreach ($prev['days'] as $k => $v) { if (!isset($d['days'][$k]) || intval($d['days'][$k]) < intval($v)) $d['days'][$k] = intval($v); }
            }
            // users (clients) wipe se bachao
            if (isset($prev['clients']) && is_array($prev['clients'])) {
                if (!isset($d['clients']) || !is_array($d['clients'])) $d['clients'] = array();
                foreach ($prev['clients'] as $k => $v) { if (!isset($d['clients'][$k])) $d['clients'][$k] = $v; }
            }
        }
    }
    // KHAALI-WIPE GUARD: bhara data 0 se replace mat karo
    $newTotal  = intval(isset($d['total'])?$d['total']:0);
    $prevScore = _data_score($prev);
    if ($newTotal === 0 && $prevScore > 0) return;      // 0 likhne se saaf inkaar

    $json = json_encode($d);
    if ($json === false || strlen($json) < 2) return;   // encode fail -> kuch mat karo

    // atomic write: temp -> rename (readers ko kabhi aadha file nahi milta)
    $tmp = $file . '.tmp.' . @getmypid() . '.' . @uniqid();
    $ok = @file_put_contents($tmp, $json, LOCK_EX);
    if ($ok !== false && $ok === strlen($json)) {
        @chmod($tmp, 0644);
        if (@rename($tmp, $file)) {
            @copy($file, $file . '.bak');               // shadow backup
        } else {
            @unlink($tmp);
        }
    } else {
        @unlink($tmp);
    }
}
// Roz ek dated backup apne aap (cron ki zaroorat nahi) — 14 rakho.
function maybe_backup($file, &$d) {
    $now = time();
    if ($now - intval(isset($d['lastBackup'])?$d['lastBackup']:0) < 86400) return;
    if (_data_score($d) <= 0) return;                 // khaali ka backup nahi
    $dir = dirname($file) . '/backups';
    if (!is_dir($dir)) @mkdir($dir, 0755, true);
    if (@file_put_contents($dir . '/stats-' . date('Y-m-d') . '.json', json_encode($d)) !== false) {
        $all = @glob($dir . '/stats-*.json');
        if ($all && count($all) > 14) { sort($all); foreach (array_slice($all, 0, count($all) - 14) as $old) @unlink($old); }
        $d['lastBackup'] = $now;
    }
}
function today_str() { return date('Y-m-d'); }
function hour_key()  { return date('Y-m-d-H'); }
function bump(&$arr, $key, $by=1) { $key=trim((string)$key); if($key==='')return; $arr[$key]=(isset($arr[$key])?intval($arr[$key]):0)+$by; }

// user ka asli IP (CDN/proxy ke peeche bhi)
function client_ip() {
    foreach (array('HTTP_CF_CONNECTING_IP','HTTP_X_FORWARDED_FOR','HTTP_X_REAL_IP','REMOTE_ADDR') as $h) {
        if (!empty($_SERVER[$h])) {
            $ip = trim(explode(',', $_SERVER[$h])[0]);
            if (filter_var($ip, FILTER_VALIDATE_IP)) return $ip;
        }
    }
    return '';
}
// IP se asli DESH (Windows-locale galat ho to bhi sahi). Pehle host/CDN header,
// warna free ip-api.com. Result caller cache karta hai (har user par ek baar).
function geo_country($ip) {
    foreach (array('HTTP_CF_IPCOUNTRY','GEOIP_COUNTRY_CODE','HTTP_X_COUNTRY_CODE') as $h) {
        if (!empty($_SERVER[$h]) && strlen($_SERVER[$h])===2 && strtoupper($_SERVER[$h])!=='XX')
            return strtoupper($_SERVER[$h]);
    }
    if ($ip==='' || $ip==='127.0.0.1' || strpos($ip,'192.168.')===0 || strpos($ip,'10.')===0 || strpos($ip,'172.16.')===0)
        return '';
    $url = "http://ip-api.com/json/" . urlencode($ip) . "?fields=countryCode";
    $s = '';
    if (function_exists('curl_init')) {
        $ch = curl_init($url);
        curl_setopt_array($ch, array(CURLOPT_RETURNTRANSFER=>true, CURLOPT_TIMEOUT=>3, CURLOPT_CONNECTTIMEOUT=>2));
        $s = @curl_exec($ch); curl_close($ch);
    } else {
        $s = @file_get_contents($url, false, stream_context_create(array('http'=>array('timeout'=>3))));
    }
    if ($s) { $j = json_decode($s, true); if (isset($j['countryCode']) && strlen($j['countryCode'])===2) return strtoupper($j['countryCode']); }
    return '';
}

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
    // DESH: pehle IP se (asli location) — ek baar geolocate karke cache.
    // Windows-locale (req['c']) sirf fallback, kyunki wo aksar galat hota hai
    // (bahut se Indian users ka Windows 'English (US)' par hota -> galat US).
    $ip = client_ip();
    if ($ip !== '' && (!isset($c['gip']) || $c['gip'] !== $ip)) {   // har IP par SIRF EK BAAR
        $gc = geo_country($ip);
        if ($gc !== '') $c['gcc'] = $gc;
        $c['gip'] = $ip;                                            // fail ho to bhi dobara try nahi
    }
    if (!empty($c['gcc'])) {
        $c['country'] = $c['gcc'];                       // IP-country authoritative
    } elseif (empty($c['country']) && !empty($req['c'])) {
        $c['country'] = substr($req['c'], 0, 4);         // fallback: app-locale
    }
    if (!empty($req['m']))  $c['method']  = substr($req['m'], 0, 10);
    if (!empty($req['u']))  $c['name']    = substr($req['u'], 0, 40);
    if (!empty($req['sm'])) $c['model']   = substr($req['sm'], 0, 40);   // scanner model
    if (!empty($req['lang'])) $c['lang']  = substr($req['lang'], 0, 4);  // hi / en
    if (!empty($req['mode'])) $c['mode']  = substr($req['mode'], 0, 8);  // simple / full
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

// Telegram par phone-notification bhejo (token+chat set ho to)
function tg_send($msg) {
    global $TELEGRAM_TOKEN, $TELEGRAM_CHAT;
    if (empty($TELEGRAM_TOKEN) || empty($TELEGRAM_CHAT)) return false;
    $url = "https://api.telegram.org/bot" . $TELEGRAM_TOKEN . "/sendMessage";
    $q = http_build_query(array('chat_id'=>$TELEGRAM_CHAT,'text'=>$msg,'parse_mode'=>'HTML','disable_web_page_preview'=>'true'));
    if (function_exists('curl_init')) {
        $ch = curl_init($url);
        curl_setopt_array($ch, array(CURLOPT_POST=>true,CURLOPT_POSTFIELDS=>$q,CURLOPT_RETURNTRANSFER=>true,CURLOPT_TIMEOUT=>8,CURLOPT_SSL_VERIFYPEER=>false));
        @curl_exec($ch); curl_close($ch);
    } else {
        @file_get_contents($url . '?' . $q);
    }
    return true;
}

// Public/safe stats — koi personal data nahi (widget + transparency page ke liye)
function public_stats($d) {
    $today = today_str(); $now = time();
    $cos = array();
    foreach ($d['clients'] as $c) { if(!empty($c['blocked']))continue; $co=trim(isset($c['country'])?$c['country']:''); if($co!=='') $cos[$co]=(isset($cos[$co])?$cos[$co]:0)+1; }
    arsort($cos);
    $online = 0; foreach ($d['online'] as $ts) { if ($now-intval($ts)<=300) $online++; }
    return array(
        'total'=>intval($d['total']),
        'today'=>intval(isset($d['days'][$today])?$d['days'][$today]:0),
        'users'=>count($d['clients']),
        'countries'=>count($cos),
        'topCountries'=>array_slice(array_keys($cos),0,6),
        'online'=>$online,
        'imports'=>intval($d['imports']),
        'prints'=>intval($d['prints']),
        'site'=>$GLOBALS['SITE_NAME'],
        'time'=>date('Y-m-d H:i')
    );
}

// =================================================================
//  CRON:  ?cron=daily&key=CRONKEY   (email report + auto-backup)
//  CLI:   php stats.php cron=daily key=CRONKEY
// =================================================================
if (PHP_SAPI === 'cli' && isset($argv)) { foreach ($argv as $a){ if(strpos($a,'=')!==false){ list($kk,$vv)=explode('=',$a,2); $_GET[$kk]=$vv; } } }
if (isset($_GET['cron'])) {
    $ckey = isset($_GET['key']) ? (string)$_GET['key'] : '';
    $ctype = $_GET['cron'];
    if (($ctype==='daily' || $ctype==='weekly') && hash_equals($CRON_KEY, $ckey)) {
        $d = load_data($DATA_FILE);
        // auto-backup (roz ek, 14 rakho)
        if (!is_dir($BACKUP_DIR)) @mkdir($BACKUP_DIR,0755,true);
        $bf = $BACKUP_DIR.'/stats-'.date('Y-m-d').'.json';
        @file_put_contents($bf, json_encode($d));
        $all=glob($BACKUP_DIR.'/stats-*.json'); if($all && count($all)>14){ sort($all); foreach(array_slice($all,0,count($all)-14) as $old) @unlink($old); }
        $d['lastBackup']=time(); save_data($DATA_FILE,$d);
        $t=today_str(); $now=time();
        if ($ctype==='weekly') {
            // pichhle 7 din ka jod
            $wk=0; for($i=1;$i<=7;$i++){ $k=date('Y-m-d',$now-$i*86400); $wk+=intval(isset($d['days'][$k])?$d['days'][$k]:0); }
            $newW=0; foreach($d['clients'] as $c){ if($now-intval($c['first'])<=7*86400)$newW++; }
            $msg="ApneScan — Weekly Report\n\nPichhle 7 din ke scans: $wk\nNaye users (7 din): $newW\n"
                ."Total (all-time): ".intval($d['total'])."\nTotal users: ".count($d['clients'])."\n\nDashboard: ?admin=...\n";
            if ($ADMIN_EMAIL!=='') @mail($ADMIN_EMAIL, "ApneScan weekly — $wk scans", $msg, "From: ApneScan <no-reply@apnesoft.com>");
            tg_send("📊 <b>ApneScan — Weekly</b>\nScans (7 din): <b>$wk</b>\nNaye users: <b>$newW</b>\nTotal: <b>".intval($d['total'])."</b>");
            header('Content-Type: text/plain'); echo "cron weekly ok\n".$msg; exit;
        }
        // daily email + telegram
        $y=date('Y-m-d',$now-86400);
        $yc=intval(isset($d['days'][$y])?$d['days'][$y]:0);
        $tc=intval(isset($d['days'][$t])?$d['days'][$t]:0);
        $newY=0; foreach($d['clients'] as $c){ if(date('Y-m-d',intval($c['first']))===$y)$newY++; }
        $ncr=count($d['crashes']);
        $msg="ApneScan — Daily Report ($y)\n\n"
            ."Kal ke scans: $yc\nAaj abhi tak: $tc\nTotal (all-time): ".intval($d['total'])."\n"
            ."Total users: ".count($d['clients'])."\nKal naye users: $newY\n"
            ."Crashes stored: ".$ncr."\nFeedback stored: ".count($d['feedback'])."\n\n"
            ."Backup: ".basename($bf)."\nDashboard: open ?admin=...\n";
        if ($ADMIN_EMAIL!=='') @mail($ADMIN_EMAIL, "ApneScan daily — $yc scans kal", $msg, "From: ApneScan <no-reply@apnesoft.com>");
        tg_send("📊 <b>ApneScan — Daily</b>\nKal: <b>$yc</b> scans · Naye users: <b>$newY</b>\nAaj tak: <b>$tc</b>\nTotal: <b>".intval($d['total'])."</b>".($ncr?"\n💥 $ncr crash report":""));
        header('Content-Type: text/plain'); echo "cron ok\n".$msg; exit;
    }
    header('Content-Type: text/plain'); echo "cron: bad key"; exit;
}

// =================================================================
//  PUBLIC WIDGET / TRANSPARENCY (login nahi chahiye — sirf safe ginti)
//    ?widget=count  -> JSON {total,today,...}   (apni website me use karo)
//    ?widget=1      -> embed-ready HTML snippet (iframe me lagao)
//    ?public=1      -> sundar public transparency page
// =================================================================
if (isset($_GET['widget'])) {
    $P = public_stats(load_data($DATA_FILE));
    if ($_GET['widget']==='count') { header('Content-Type: application/json'); echo json_encode($P); exit; }
    header('Content-Type: text/html; charset=utf-8');
    ?><!doctype html><meta charset="utf-8"><style>
    *{margin:0;box-sizing:border-box}body{font-family:Inter,system-ui,Segoe UI,Arial;background:transparent}
    .w{display:inline-flex;gap:14px;align-items:center;background:linear-gradient(120deg,#12325f,#178a8a);color:#fff;padding:12px 18px;border-radius:14px;box-shadow:0 6px 20px rgba(9,20,45,.25)}
    .w .n{font-size:24px;font-weight:800;line-height:1}.w .l{font-size:10px;opacity:.85;margin-top:2px}
    .w .d{width:1px;height:34px;background:rgba(255,255,255,.25)}
    </style><div class="w"><div>📊</div><div><div class="n" id="t">…</div><div class="l"><?php echo htmlspecialchars($P['site']); ?> — Total scans worldwide</div></div>
    <div class="d"></div><div><div class="n" id="d">…</div><div class="l">Aaj</div></div></div>
    <script>var P=<?php echo json_encode($P); ?>;function f(n){return (n||0).toLocaleString();}document.getElementById('t').textContent=f(P.total);document.getElementById('d').textContent=f(P.today);</script>
    <?php exit;
}
if (isset($_GET['public'])) {
    $P = public_stats(load_data($DATA_FILE));
    header('Content-Type: text/html; charset=utf-8');
    function _fl($cc){ if(strlen($cc)!==2) return '🏳'; $a=strtoupper($cc); return mb_convert_encoding('&#'.(127397+ord($a[0])).';&#'.(127397+ord($a[1])).';','UTF-8','HTML-ENTITIES'); }
    ?><!doctype html><html lang="hi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
    <title><?php echo htmlspecialchars($P['site']); ?> — Live Stats</title><style>
    *{margin:0;box-sizing:border-box}body{font-family:Inter,system-ui,Segoe UI,Arial;background:#0a0f1c;color:#e2e8f0;min-height:100vh;display:flex;align-items:center;justify-content:center;padding:24px}
    .box{max-width:760px;width:100%;text-align:center}
    h1{font-size:26px;margin-bottom:6px}.sub{color:#8ea0bd;font-size:13px;margin-bottom:26px}
    .g{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:14px}
    .c{background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.12);border-radius:16px;padding:20px}
    .c .n{font-size:30px;font-weight:800;background:linear-gradient(120deg,#3987e5,#2dd4bf);-webkit-background-clip:text;background-clip:text;color:transparent}
    .c .l{color:#a9b6ce;font-size:12px;margin-top:6px}
    .f{color:#7f8fab;font-size:11px;margin-top:22px}.cos{margin-top:16px;font-size:20px;letter-spacing:4px}
    </style></head><body><div class="box">
    <h1>🌍 <?php echo htmlspecialchars($P['site']); ?> — Worldwide</h1>
    <div class="sub">100% FREE document scanner · live ginti (koi personal data nahi)</div>
    <div class="g">
      <div class="c"><div class="n" id="total">…</div><div class="l">📄 Total scans</div></div>
      <div class="c"><div class="n" id="today">…</div><div class="l">📅 Aaj ke scans</div></div>
      <div class="c"><div class="n" id="users">…</div><div class="l">👥 Total users</div></div>
      <div class="c"><div class="n" id="countries">…</div><div class="l">🌍 Countries</div></div>
      <div class="c"><div class="n" id="online">…</div><div class="l">🟢 Abhi online</div></div>
      <div class="c"><div class="n" id="paper">…</div><div class="l">🌿 Paper digitized</div></div>
    </div>
    <div class="cos"><?php foreach($P['topCountries'] as $cc) echo _fl($cc).' '; ?></div>
    <div class="f">Updated <?php echo htmlspecialchars($P['time']); ?> · apnescan.apnesoft.com</div>
    </div><script>var P=<?php echo json_encode($P); ?>;function f(n){return (n||0).toLocaleString();}
    ['total','today','users','countries','online'].forEach(function(k){document.getElementById(k).textContent=f(P[k]);});
    document.getElementById('paper').textContent=f(P.total);
    setTimeout(function(){location.reload();},60000);</script></body></html><?php exit;
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
        save_data($DATA_FILE,$d,true);   // admin ka jaan-boojhkar action (purge/clear) -> force
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
    // ================= DEEP WORLDWIDE ANALYTICS =================
    // day-of-week (Mon..Sun)
    $dow=array_fill(0,7,0);
    foreach($d['days'] as $k=>$v){ $ts=strtotime($k); if($ts){ $dow[intval(date('w',$ts))]+=intval($v); } }
    $S['dow']=array($dow[1],$dow[2],$dow[3],$dow[4],$dow[5],$dow[6],$dow[0]);
    $S['weekdayTotal']=$dow[1]+$dow[2]+$dow[3]+$dow[4]+$dow[5];
    $S['weekendTotal']=$dow[6]+$dow[0];
    // month-by-month (last 12)
    $mbm=array();
    for($i=11;$i>=0;$i--){ $mk=date('Y-m',strtotime("first day of -$i month",$now)); $sum=0; foreach($d['days'] as $k=>$v){ if(strpos($k,$mk.'-')===0)$sum+=intval($v); } $mbm[]=array(date('M y',strtotime($mk.'-01')),$sum); }
    $S['monthByMonth']=$mbm;
    // 7x24 heatmap
    $hw=array(); for($i=0;$i<7;$i++)$hw[$i]=array_fill(0,24,0);
    foreach($d['hours'] as $hk=>$hv){ if(preg_match('/^(\d{4}-\d{2}-\d{2})-(\d{1,2})$/',$hk,$m)){ $ts=strtotime($m[1]); if($ts){ $hh=intval($m[2]); if($hh>=0&&$hh<24)$hw[intval(date('w',$ts))][$hh]+=intval($hv); } } }
    $S['heat7x24']=array($hw[1],$hw[2],$hw[3],$hw[4],$hw[5],$hw[6],$hw[0]);
    // week-over-week
    $tw=0; for($i=0;$i<7;$i++){ $k=date('Y-m-d',$now-$i*86400); $tw+=intval(isset($d['days'][$k])?$d['days'][$k]:0); }
    $lw=0; for($i=7;$i<14;$i++){ $k=date('Y-m-d',$now-$i*86400); $lw+=intval(isset($d['days'][$k])?$d['days'][$k]:0); }
    $S['thisWeek7']=$tw; $S['lastWeek7']=$lw; $S['wowPct']=($lw>0)?round(100*($tw-$lw)/$lw):($tw>0?100:0);
    // year-over-year
    $ty=0;$ly=0; $cy=date('Y'); $py=$cy-1;
    foreach($d['days'] as $k=>$v){ if(strpos($k,$cy.'-')===0)$ty+=intval($v); elseif(strpos($k,$py.'-')===0)$ly+=intval($v); }
    $S['thisYear']=$ty; $S['lastYear']=$ly;
    // country depth
    $usersByCo=array();
    foreach($d['clients'] as $c){ if(!empty($c['blocked']))continue; $co=trim($c['country']); if($co!=='')bump($usersByCo,$co,1); }
    $S['usersByCountry']=$usersByCo;
    $perUserCo=array(); foreach($usersByCo as $co=>$u){ $sc=isset($S['scansByCountry'][$co])?$S['scansByCountry'][$co]:0; $perUserCo[$co]=$u?round($sc/$u,1):0; }
    $S['perUserByCountry']=$perUserCo;
    $moStart=strtotime(date('Y-m-01')); $coFirst=array();
    foreach($d['clients'] as $c){ if(!empty($c['blocked']))continue; $co=trim($c['country']); if($co==='')continue; $f=intval($c['first']); if(!isset($coFirst[$co])||$f<$coFirst[$co])$coFirst[$co]=$f; }
    $newCos=array(); foreach($coFirst as $co=>$f){ if($f>=$moStart)$newCos[]=$co; } $S['newCountries']=$newCos;
    // segments
    $seg=array('new'=>0,'regular'=>0,'power'=>0,'asleep'=>0);
    foreach($d['clients'] as $c){ if(!empty($c['blocked']))continue; $sc=intval($c['scans']); $last=intval($c['last']); $first=intval($c['first']);
      if($now-$last>14*86400)$seg['asleep']++; elseif($sc>=50)$seg['power']++; elseif($now-$first<7*86400)$seg['new']++; else $seg['regular']++; }
    $S['segments']=$seg;
    $S['avgPerUser']=count($d['clients'])?round($S['total']/max(1,count($d['clients'])),1):0;
    // distribution
    $bk=array('1-5'=>0,'6-20'=>0,'21-50'=>0,'51-100'=>0,'100+'=>0);
    foreach($d['clients'] as $c){ if(!empty($c['blocked']))continue; $sc=intval($c['scans']); if($sc<=0)continue; if($sc<=5)$bk['1-5']++; elseif($sc<=20)$bk['6-20']++; elseif($sc<=50)$bk['21-50']++; elseif($sc<=100)$bk['51-100']++; else $bk['100+']++; }
    $S['scanDist']=$bk;
    // retention curve
    $rcv=array(); foreach(array(1,3,7,14,30) as $dd){ $e=0;$r=0;
      foreach($d['clients'] as $c){ if(!empty($c['blocked']))continue; $first=intval($c['first']); if(!$first||$now-$first<$dd*86400)continue; $e++; $act=isset($c['active'])?$c['active']:array(); $cut=date('Y-m-d',$first+$dd*86400); foreach($act as $ad){ if($ad>=$cut){$r++;break;} } }
      $rcv[]=array($dd,$e?round(100*$r/$e):0); }
    $S['retCurve']=$rcv;
    $S['stickiness']=$S['mau']?round(100*$S['dau']/$S['mau']):0;
    $rep=0;$tot=0; foreach($d['clients'] as $c){ if(!empty($c['blocked']))continue; $tot++; if(count(isset($c['active'])?$c['active']:array())>=2)$rep++; }
    $S['repeatRate']=$tot?round(100*$rep/$tot):0;
    // scanner brands
    $brands=array(); foreach($d['scanners'] as $mm=>$cnt){ $bw=trim(strtok($mm,' ')); if($bw!=='')bump($brands,$bw,$cnt); } $S['brands']=$brands;
    // avg pages/scan
    $S['scanEvents']=intval(isset($d['scanEvents'])?$d['scanEvents']:0);
    $S['avgPages']=$S['scanEvents']?round($S['total']/max(1,$S['scanEvents']),1):0;
    // most common settings
    $tk=function($a){ if(!is_array($a)||!$a)return array('—',0); arsort($a); $k=key($a); return array((string)$k,intval($a[$k])); };
    $S['topDpi']=$tk($d['dpis']); $S['topColor']=$tk($d['colors']); $S['topSize']=$tk($d['sizes']);
    // growth + churn rate
    $ntm=0;$nlm=0; $lmS=strtotime(date('Y-m-01',strtotime('-1 month'))); $lmE=$moStart-1;
    foreach($d['clients'] as $c){ if(!empty($c['blocked']))continue; $f=intval($c['first']); if($f>=$moStart)$ntm++; elseif($f>=$lmS&&$f<=$lmE)$nlm++; }
    $S['newThisMonth']=$ntm; $S['newLastMonth']=$nlm; $S['userGrowthPct']=$nlm?round(100*($ntm-$nlm)/$nlm):($ntm?100:0);
    $chn=0; foreach($d['clients'] as $c){ if(!empty($c['blocked']))continue; if($now-intval($c['last'])>=14*86400 && intval($c['scans'])>=1)$chn++; }
    $S['churnRate']=count($d['clients'])?round(100*$chn/max(1,count($d['clients']))):0; $S['churnCount']=$chn;
    $S['netGrowthMo']=$ntm-$chn;
    // milestone countdown
    $miles=array(100,500,1000,5000,10000,25000,50000,100000,250000,500000,1000000);
    $next=null; foreach($miles as $m){ if($S['total']<$m){$next=$m;break;} }
    $prevM=0; foreach($miles as $m){ if($m<=$S['total'])$prevM=$m; }
    $S['nextMilestone']=$next; $S['milestoneLeft']=$next?($next-$S['total']):0; $S['milestonePct']=$next?round(100*($S['total']-$prevM)/max(1,($next-$prevM))):100;
    // global streak
    $streak=0; for($i=0;$i<400;$i++){ $k=date('Y-m-d',$now-$i*86400); if(intval(isset($d['days'][$k])?$d['days'][$k]:0)>0)$streak++; elseif($i>0)break; } $S['globalStreak']=$streak;
    // forecast + projection
    $S['forecastMonth']=round($S['dailyAvg']*30);
    $dLeft=(intval(date('z',mktime(0,0,0,12,31,intval($cy))))-intval(date('z')));
    $S['yearEndProj']=$S['thisYear']+round($S['dailyAvg']*max(0,$dLeft));
    // live feed + last hour/24h
    $S['recentScans']=array_reverse(array_slice(isset($d['recentScans'])?$d['recentScans']:array(),-20));
    $S['lastHour']=intval(isset($d['hours'][hour_key()])?$d['hours'][hour_key()]:0);
    $l24=0; for($i=0;$i<24;$i++){ $hk=date('Y-m-d-H',$now-$i*3600); $l24+=intval(isset($d['hours'][$hk])?$d['hours'][$hk]:0); } $S['last24h']=$l24;

    // ================= TOOLS & IMPACT =================
    $totU=max(1,count($d['clients']));
    $feats=isset($d['features'])?$d['features']:array();
    // feature adoption %
    $adopt=array(); foreach((isset($d['featUsers'])?$d['featUsers']:array()) as $f=>$us){ $adopt[$f]=round(100*count($us)/$totU); }
    $S['featAdoption']=$adopt;
    // share pie
    $S['sharePie']=array('whatsapp'=>intval(isset($feats['whatsapp'])?$feats['whatsapp']:0),'email'=>intval(isset($feats['email'])?$feats['email']:0),'print'=>intval($d['prints']));
    // metrics
    $met=isset($d['metrics'])?$d['metrics']:array();
    $S['mbSaved']=round(intval(isset($met['kbSaved'])?$met['kbSaved']:0)/1024,1);
    $S['mergePages']=intval(isset($met['pg_merge'])?$met['pg_merge']:0);
    // multi-feature users
    $mf=array('0'=>0,'1'=>0,'2'=>0,'3'=>0,'4+'=>0);
    foreach($d['clients'] as $c){ if(!empty($c['blocked']))continue; $nn=count(isset($c['feats'])?$c['feats']:array()); if($nn<=0)$mf['0']++; elseif($nn==1)$mf['1']++; elseif($nn==2)$mf['2']++; elseif($nn==3)$mf['3']++; else $mf['4+']++; }
    $S['multiFeature']=$mf;
    // preferences
    $langD=array(); $modeD=array();
    foreach($d['clients'] as $c){ if(!empty($c['blocked']))continue; $l=trim(isset($c['lang'])?$c['lang']:''); if($l!=='')bump($langD,$l,1); $mo=trim(isset($c['mode'])?$c['mode']:''); if($mo!=='')bump($modeD,$mo,1); }
    $S['langDist']=$langD; $S['modeDist']=$modeD;
    // IMPACT
    $paperSheets=$S['total']+intval($d['imports']);
    $S['impactDocs']=$S['scanEvents']+intval($d['imports']);
    $S['impactPaper']=$paperSheets;
    $S['impactDataMB']=$S['mbSaved'];
    $S['impactTrees']=round($paperSheets/8333,2);
    $shr=intval(isset($feats['whatsapp'])?$feats['whatsapp']:0)+intval(isset($feats['email'])?$feats['email']:0);
    $S['impactShares']=$shr;
    $S['impactMinutes']=round(($S['total']+intval($d['imports']))*0.5 + $shr*2);
    $S['impactHours']=round($S['impactMinutes']/60);

    // ================= SMART SUGGESTIONS (aage kya banayein) =================
    $sug=array();
    $addsug=function($p,$ic,$t,$dd) use (&$sug){ $sug[]=array('p'=>$p,'ic'=>$ic,'t'=>$t,'d'=>$dd); };
    $ncrash=count(isset($d['crashes'])?$d['crashes']:array());
    if ($ncrash>0){
        $cv=array(); foreach($d['crashes'] as $c){ $vv=trim(isset($c['v'])?$c['v']:''); if($vv!=='')bump($cv,$vv,1); }
        arsort($cv); $topcv=$cv?key($cv):'';
        $addsug(1,'💥','App crashes theek karein (sabse zaroori)', $ncrash.' crash report aaye hain'.($topcv?' — sabse zyada v'.$topcv.' me':'').'. "Feedback & System" tab me detail dekhein aur agli update me fix karein.');
    }
    if ($S['oldVersionPct']>=30){
        $addsug(1,'🔔','Purane version wale users ko update karayein', $S['oldVersionPct'].'% users purane version par hain. "Overview → Broadcast" se sabhi ya sirf purane-version users ko update ka message bhejein.');
    }
    if (isset($S['ratingCount']) && $S['ratingCount']>0 && $S['avgRating']>0 && $S['avgRating']<4){
        $addsug(1,'⭐','Rating kam hai — feedback par kaam karein', 'Average rating '.$S['avgRating'].'⭐ hai. Users ke feedback padhkar sabse common shikayat sabse pehle theek karein.');
    }
    $nfb=count(isset($d['feedback'])?$d['feedback']:array());
    if ($nfb>0){
        $addsug(2,'💬','Users ke '.$nfb.' feedback padhein', 'Feedback me aksar nayi feature ki demand chhipi hoti hai. "Feedback & System" tab me padhkar sabse zyada maangi cheez banayein.');
    }
    if (!empty($S['featAdoption'])){
        $la=array(); foreach($S['featAdoption'] as $f=>$p){ if($p>0 && $p<15) $la[$f]=$p; }
        asort($la); $i=0; foreach($la as $f=>$p){ if($i++>=2)break;
            $addsug(2,'🧰','‘'.ucfirst($f).'’ feature kam use hota', 'Sirf '.$p.'% users ‘'.$f.'’ use karte hain. Ise aur aasan banayein, button saamne layein, ya ek chhota tutorial dein.'); }
    }
    if ($S['churnRate']>=30 && $S['users']>=5){
        $addsug(2,'😴','Chhute hue users wapas laayein', $S['churnRate'].'% users 14+ din se gayab hain. Reminder/notification ya koi nayi feature se wapas laane ka socho.');
    }
    if (isset($S['ret7']) && $S['ret7']<30 && $S['users']>=5){
        $addsug(2,'🔁','Pehle-hafte ka retention kam', 'Week-1 me sirf '.$S['ret7'].'% users wapas aate. Pehli scan ko aur aasan/tez banayein aur ek welcome-guide dein.');
    }
    if ($S['oneTime']>0 && $S['oneTime']>($S['powerUsers']*2) && $S['users']>=6){
        $addsug(2,'1️⃣','Bahut se ek-baar aane wale users', $S['oneTime'].' log sirf ek baar aaye. Pehli baar ka experience (setup + pehli scan) aur simple/aasan banayein.');
    }
    if (isset($S['userGrowthPct'])){
        if ($S['userGrowthPct']<0) $addsug(2,'📉','Naye users ghat rahe', 'Is mahine naye users '.$S['userGrowthPct'].'% (pichhle mahine se kam). Share/referral feature ya thoda marketing push karein.');
        elseif ($S['userGrowthPct']>=50) $addsug(3,'🚀','Users tezi se badh rahe 🎉', '+'.$S['userGrowthPct'].'% growth! Server/scale aur support ka dhyaan rakhein, momentum banaye rakhein.');
    }
    if (!empty($S['features'])){
        $ff=$S['features']; arsort($ff); $tf=key($ff);
        if ($tf && $ff[$tf]>0) $addsug(3,'🌟','Sabse popular feature: ‘'.ucfirst($tf).'’', 'Log ‘'.$tf.'’ sabse zyada use karte. Isme aur options/power add karein — yahi aapki taakat hai.');
    }
    if (!empty($S['scanners'])){
        $scn=$S['scanners']; arsort($scn); $ts=key($scn);
        if ($ts!=='') $addsug(3,'🖨','‘'.$ts.'’ scanner sabse zyada', 'Zyadatar log ‘'.$ts.'’ use karte. Uske liye khaas test/optimize karein taaki us par sab kuch perfect chale.');
    }
    if (isset($S['topDpi'][0]) && $S['topDpi'][0]!=='—'){
        $addsug(3,'🎚','Sabse common setting: '.$S['topDpi'][0].' DPI · '.(isset($S['topColor'][0])?$S['topColor'][0]:'—'), 'Default aur scan-speed isi setting ke liye optimize karein (zyadatar log yahi use karte).');
    }
    if (!empty($S['langDist'])){
        $ld=$S['langDist']; $tot=array_sum($ld);
        if ($tot>0 && isset($ld['hi']) && $ld['hi']*2>=$tot) $addsug(3,'🇮🇳','Zyadatar users Hindi me', 'Hindi tutorial/help video aur Hindi-friendly features banayein — yahi aapke asli users hain.');
    }
    if (!empty($S['modeDist'])){
        $md=$S['modeDist']; $tot=array_sum($md);
        if ($tot>0 && isset($md['simple']) && $md['simple']*2>=$tot) $addsug(3,'🟢','Log Simple mode pasand karte', 'Simple mode ko aur polish karein aur naye users ke liye ise default rakhein.');
    }
    if (!empty($S['nextMilestone']) && $S['milestoneLeft']<=($S['nextMilestone']*0.1)){
        $addsug(3,'🎯','Milestone kareeb: '.number_format($S['nextMilestone']), 'Bas '.number_format($S['milestoneLeft']).' scans aur! Paar hote hi app me celebrate/announce karein (Broadcast).');
    }
    if (empty($sug)){
        $addsug(3,'📊','Abhi data thoda kam hai', 'Jaise-jaise zyada users naye version (v136+) par aayenge, yahan aapko "aage kya banayein" ke smart suggestions milne lagenge.');
    }
    usort($sug, function($a,$b){ return $a['p']-$b['p']; });
    $S['suggestions']=$sug;

    // ---- AI-style daily summary (Hindi) — rozana ek nazar ----
    $ai=array();
    $line="Aaj ab tak <b>".number_format($S['today'])."</b> scans hue";
    if ($S['yesterday']>0){ $dp=round(100*($S['today']-$S['yesterday'])/$S['yesterday']); $line.=" (kal se ".($dp>=0?"+$dp% 📈":"$dp% 📉").")"; }
    $ai[]=$line;
    $ai[]="Ab tak kul <b>".number_format($S['total'])."</b> scans aur <b>".number_format($S['users'])."</b> users";
    if ($S['newToday']>0) $ai[]="Aaj <b>".$S['newToday']."</b> naye users jude 🎉";
    if (!empty($S['features'])){ $ff=$S['features']; arsort($ff); $ai[]="Sabse zyada ‘<b>".ucfirst(key($ff))."</b>’ feature chala"; }
    $ncrash2=count(isset($d['crashes'])?$d['crashes']:array());
    $ai[]=$ncrash2>0 ? "⚠️ <b>$ncrash2</b> crash report aaye — dekhna zaroori" : "✅ Koi crash nahi, sab theek";
    if ($S['online']>0) $ai[]="Abhi <b>".$S['online']."</b> log online";
    if (!empty($S['nextMilestone'])) $ai[]="Agle milestone (".number_format($S['nextMilestone']).") tak bas <b>".number_format($S['milestoneLeft'])."</b> baaki 🎯";
    $S['aiSummary']=implode('. ',$ai).'.';

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
  body{margin:0;font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:var(--bg);color:var(--fg);-webkit-font-smoothing:antialiased;font-size:11px;line-height:1.45}
  a{color:inherit}
  /* app bar */
  header{background:var(--hd);color:#fff;padding:11px 18px;position:sticky;top:0;z-index:20;box-shadow:0 3px 14px rgba(9,20,45,.25);display:flex;align-items:center;gap:12px;flex-wrap:wrap}
  header .brand{display:flex;align-items:center;gap:9px;font-size:15px;font-weight:800;letter-spacing:-.01em}
  header .brand .logo{width:31px;height:31px;border-radius:9px;background:rgba(255,255,255,.16);display:flex;align-items:center;justify-content:center;font-size:16px}
  header .brand small{display:block;font-size:9.5px;font-weight:500;opacity:.8;letter-spacing:.02em}
  header .live{display:flex;align-items:center;gap:6px;background:rgba(255,255,255,.14);padding:5px 10px;border-radius:30px;font-size:10.5px}
  header .live .dot{width:7px;height:7px;border-radius:50%;background:#4ade80;box-shadow:0 0 0 0 rgba(74,222,128,.6);animation:pulse 2s infinite}
  @keyframes pulse{0%{box-shadow:0 0 0 0 rgba(74,222,128,.55)}70%{box-shadow:0 0 0 6px rgba(74,222,128,0)}100%{box-shadow:0 0 0 0 rgba(74,222,128,0)}}
  header .toolbar{margin-left:auto;display:flex;gap:6px;align-items:center}
  .iconbtn{width:32px;height:32px;border-radius:9px;border:none;background:rgba(255,255,255,.14);color:#fff;cursor:pointer;font-size:14px;display:inline-flex;align-items:center;justify-content:center;text-decoration:none;transition:background .15s,transform .1s}
  .iconbtn:hover{background:rgba(255,255,255,.28)} .iconbtn:active{transform:scale(.93)}
  .iconbtn.logout{width:auto;padding:0 11px;gap:5px;font-size:11px;font-weight:700}
  /* tab nav */
  .tabs{position:sticky;top:0;z-index:19;display:flex;gap:4px;overflow-x:auto;padding:8px 16px;
    background:var(--card);border-bottom:1px solid var(--line);box-shadow:0 2px 8px rgba(16,24,40,.05)}
  .tabs::-webkit-scrollbar{height:0}
  .tab{white-space:nowrap;border:1px solid transparent;background:transparent;color:var(--mut);
    font-size:11.5px;font-weight:700;padding:7px 13px;border-radius:9px;cursor:pointer;transition:all .14s}
  .tab:hover{background:var(--card2);color:var(--fg)}
  .tab.active{background:var(--accent);color:#fff;box-shadow:var(--sh)}
  .wrap{max-width:none;margin:14px auto;padding:0 18px}
  /* section titles */
  .sec{display:flex;align-items:center;gap:10px;margin:22px 2px 11px;font-size:10px;font-weight:800;letter-spacing:.10em;text-transform:uppercase;color:var(--mut)}
  .sec .em{font-size:12px}
  .sec::after{content:"";flex:1;height:1px;background:linear-gradient(90deg,var(--line),transparent)}
  .sec:first-of-type{margin-top:4px}
  /* KPI */
  .kpis{display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:10px}
  .kpi{position:relative;overflow:hidden;display:flex;align-items:center;gap:11px;background:var(--card);
    border:1px solid var(--line);border-radius:13px;padding:11px 12px;box-shadow:var(--sh);
    transition:transform .18s ease,box-shadow .18s ease;--kc:rgba(42,120,214,.13);--ig:linear-gradient(135deg,#dcebfe,#c2ddfb)}
  .kpi::before{content:"";position:absolute;top:-40%;right:-25%;width:100px;height:100px;border-radius:50%;
    background:radial-gradient(circle,var(--kc),transparent 68%);pointer-events:none}
  .kpi:hover{transform:translateY(-3px);box-shadow:var(--sh2)}
  .kpi .ic{position:relative;z-index:1;width:38px;height:38px;flex:none;border-radius:11px;display:flex;
    align-items:center;justify-content:center;font-size:17px;background:var(--ig);
    box-shadow:0 3px 9px var(--kc),inset 0 1px 0 rgba(255,255,255,.5)}
  .kpi .tx{position:relative;z-index:1;min-width:0}
  .kpi .n{font-size:18px;font-weight:800;line-height:1;letter-spacing:-.02em}
  .kpi .l{color:var(--mut);font-size:9.5px;margin-top:4px;font-weight:600;line-height:1.2;letter-spacing:.01em}
  .kpi.g{--kc:rgba(27,175,122,.16);--ig:linear-gradient(135deg,#d2f6e6,#a7ecca)} .kpi.g .n{color:#0f9d6b}
  .kpi.r{--kc:rgba(227,73,72,.16);--ig:linear-gradient(135deg,#fde0e0,#f9c3c3)} .kpi.r .n{color:#e34948}
  .kpi.p{--kc:rgba(108,92,240,.18);--ig:linear-gradient(135deg,#e7e2ff,#d3cbff)} .kpi.p .n{color:#6b5cf0}
  .kpi.o{--kc:rgba(235,104,52,.16);--ig:linear-gradient(135deg,#ffe4d6,#ffccb3)} .kpi.o .n{color:#e06a38}
  .kpi.y{--kc:rgba(237,161,0,.18);--ig:linear-gradient(135deg,#fdefc0,#fbe08a)} .kpi.y .n{color:#c98500}
  html[data-th=dark] .kpi .ic{background:var(--kc);box-shadow:inset 0 0 0 1px rgba(255,255,255,.06)}
  /* layout */
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
  .grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px}
  @media(max-width:900px){.grid,.grid3{grid-template-columns:1fr}}
  .card{background:var(--card);border:1px solid var(--line);border-radius:13px;padding:12px 13px;margin-bottom:12px;box-shadow:var(--sh)}
  .card h3{margin:0 0 9px;font-size:11px;font-weight:700;display:flex;align-items:center;gap:7px;color:var(--fg)}
  .card h3 .em{font-size:13px}
  .card canvas{height:150px!important;width:100%!important}
  table{width:100%;border-collapse:collapse;font-size:10.5px}
  td,th{padding:4px 6px;border-bottom:1px solid var(--line);text-align:left}
  tbody tr:last-child td{border-bottom:none}
  th{cursor:pointer;color:var(--mut);font-size:9.5px;text-transform:uppercase;letter-spacing:.04em;font-weight:700}
  .bar{height:8px;background:linear-gradient(90deg,var(--accent),#5fa0ec);border-radius:5px;min-width:5px;box-shadow:inset 0 -1px 2px rgba(0,0,0,.08)}
  /* buttons */
  .rbtn{background:var(--accent);color:#fff;border:none;border-radius:8px;padding:5px 10px;cursor:pointer;font-weight:700;margin-left:5px;text-decoration:none;display:inline-block;font-size:10.5px;box-shadow:var(--sh);transition:filter .15s} .rbtn:hover{filter:brightness(1.07)}
  .rbtn.d{background:var(--card2);color:var(--accent);border:1px solid var(--line)}
  input,select,textarea{padding:6px 9px;border:1px solid var(--line);border-radius:8px;font-size:11px;background:var(--card2);color:var(--fg);font-family:inherit;outline:none;transition:border-color .15s,box-shadow .15s}
  input:focus,select:focus,textarea:focus{border-color:var(--accent);box-shadow:0 0 0 3px rgba(42,120,214,.15)}
  .btn{background:var(--accent);color:#fff;border:none;border-radius:8px;padding:6px 11px;cursor:pointer;font-weight:700;font-size:11px;box-shadow:var(--sh);transition:filter .15s} .btn:hover{filter:brightness(1.07)}
  .btn.gray{background:#64748b}.btn.red{background:#e34948}
  .banner{background:linear-gradient(90deg,#fff7d6,#fef3c7);border:1px solid #fadf8a;color:#8a5a00;border-radius:10px;padding:9px 13px;margin-bottom:12px;font-size:11px;font-weight:600;display:none;box-shadow:var(--sh)}
  html[data-th=dark] .banner{background:linear-gradient(90deg,#3a2f10,#33290c);border-color:#5a4718;color:#fcd34d}
  .btns{margin-bottom:8px;display:flex;gap:2px;flex-wrap:wrap;align-items:center}
  .foot{color:var(--mut);font-size:10px;text-align:center;margin:18px 0 8px}
  .tag{display:inline-block;background:rgba(42,120,214,.12);color:var(--accent);border-radius:20px;padding:1px 8px;font-size:9px;margin-right:3px;font-weight:700}
  .heat{display:grid;grid-template-columns:repeat(24,1fr);gap:3px}
  .heat div{height:26px;border-radius:5px;font-size:8px;color:#fff;text-align:center;line-height:26px;font-weight:600}
  @media print{.toolbar,.btns,.no-print{display:none}.card{break-inside:avoid;box-shadow:none}.page{display:block!important}}
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
<nav class="tabs no-print" id="tabs">
  <button class="tab" data-p="overview">📊 Overview</button>
  <button class="tab" data-p="ideas">💡 Suggestions</button>
  <button class="tab" data-p="trends">📈 Trends</button>
  <button class="tab" data-p="growth">🔁 Growth &amp; Analytics</button>
  <button class="tab" data-p="tools">🧰 Tools &amp; Impact</button>
  <button class="tab" data-p="users">👤 Users</button>
  <button class="tab" data-p="devices">🌍 Activity &amp; Devices</button>
  <button class="tab" data-p="system">💬 Feedback &amp; System</button>
</nav>
<div class="wrap">
  <div id="banner" class="banner"></div>

  <div class="page" data-p="overview">
  <div class="card" style="background:linear-gradient(120deg,rgba(42,120,214,.10),rgba(23,138,138,.06));border-color:rgba(42,120,214,.25)">
    <h3><span class="em">🤖</span> Aaj ka summary (auto)</h3>
    <div id="aisum" style="font-size:12.5px;line-height:1.7;color:var(--fg)"></div>
  </div>
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

  </div><!-- /overview -->

  <div class="page" data-p="ideas">
  <div class="sec"><span class="em">💡</span> Suggestions — aage kya banayein / sudhaarein</div>
  <div class="card" style="background:linear-gradient(120deg,rgba(42,120,214,.06),transparent)"><h3><span class="em">🤖</span> Aapke worldwide data se apne-aap bane sujhav</h3>
    <div style="color:var(--mut);font-size:11px">Ye suggestions aapke asli users ke istemaal, crashes, feedback aur version ke hisaab se banaye gaye hain — priority ke saath (🔴 zaroori → 🔵 sujhav).</div>
  </div>
  <div id="ideas"></div>
  </div><!-- /ideas -->

  <div class="page" data-p="trends">
  <div class="sec"><span class="em">📈</span> Trends</div>
  <div class="grid">
    <div class="card"><h3><span class="em">📊</span> Last 7 days</h3><canvas id="wk" height="150"></canvas></div>
    <div class="card"><h3><span class="em">📈</span> Last 30 days</h3><canvas id="mo" height="150"></canvas></div>
    <div class="card"><h3><span class="em">🌍</span> Today — 24 hours</h3><canvas id="hr" height="150"></canvas></div>
    <div class="card"><h3><span class="em">🏆</span> Top users (naam ke saath)</h3><canvas id="tp" height="150"></canvas></div>
    <div class="card"><h3><span class="em">🌱</span> New users — 30 days</h3><canvas id="nu" height="150"></canvas></div>
    <div class="card"><h3><span class="em">👥</span> Total users (growth)</h3><canvas id="cu" height="150"></canvas></div>
  </div>

  </div><!-- /trends -->

  <div class="page" data-p="growth">
  <div class="sec"><span class="em">🔁</span> Growth &amp; Retention</div>
  <!-- retention + funnel + cohort -->
  <div class="grid3">
    <div class="card"><h3><span class="em">🔁</span> Retention (wapas aaye)</h3><div id="ret"></div></div>
    <div class="card"><h3><span class="em">🫗</span> Funnel</h3><div id="fun"></div></div>
    <div class="card"><h3><span class="em">📅</span> Weekly cohorts (30-din tak tike)</h3><div id="coh"></div></div>
  </div>

  <div class="card"><h3><span class="em">🕐</span> Busiest hours (all-time)</h3><div class="heat" id="heat"></div></div>

  <div class="sec"><span class="em">🌐</span> Deep Worldwide Analytics</div>
  <div class="kpis" id="kpis2" style="margin-bottom:16px"></div>

  <!-- milestone + forecast + live feed -->
  <div class="grid3">
    <div class="card"><h3><span class="em">🎯</span> Agle milestone tak</h3><div id="mile"></div></div>
    <div class="card"><h3><span class="em">🔮</span> Forecast &amp; projection</h3><div id="fore"></div></div>
    <div class="card"><h3><span class="em">🔴</span> Live scan feed</h3><div id="feed" style="max-height:200px;overflow:auto"></div></div>
  </div>

  <div class="grid">
    <div class="card"><h3><span class="em">📅</span> Din-vaar (Mon–Sun)</h3><canvas id="dowc" height="150"></canvas></div>
    <div class="card"><h3><span class="em">📆</span> Mahina-dar-mahina (12 month)</h3><canvas id="mbmc" height="150"></canvas></div>
    <div class="card"><h3><span class="em">🥧</span> User segments</h3><canvas id="segc" height="150"></canvas></div>
    <div class="card"><h3><span class="em">📊</span> Scans distribution (per user)</h3><canvas id="distc" height="150"></canvas></div>
    <div class="card"><h3><span class="em">📉</span> Retention curve</h3><canvas id="retcv" height="150"></canvas></div>
    <div class="card"><h3><span class="em">🏭</span> Scanner brands (market-share)</h3><div id="brands"></div></div>
  </div>

  <div class="card"><h3><span class="em">🗓</span> World rhythm — din × ghanta (7×24)</h3>
    <div style="overflow-x:auto"><div id="heat7"></div></div>
  </div>

  <div class="card"><h3><span class="em">🗺</span> World map — scans by country</h3><canvas id="wmap" height="300" style="height:340px!important"></canvas></div>

  <div class="grid">
    <div class="card"><h3><span class="em">🌍</span> Country depth (users · scans · per-user · %)</h3><div style="max-height:320px;overflow:auto"><div id="coDeep"></div></div></div>
    <div class="card"><h3><span class="em">🎚</span> Sabse common setting</h3><div id="common"></div>
      <h3 style="margin-top:14px"><span class="em">⚡</span> Weekday vs Weekend</h3><div id="wwe"></div></div>
  </div>

  <!-- compare tool -->
  <div class="card no-print"><h3><span class="em">⚖️</span> Compare tool</h3>
    <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:10px">
      <select id="cmpType"><option value="country">Desh vs Desh</option><option value="version">Version vs Version</option><option value="date">Do date-range</option></select>
      <span id="cmpAB"></span>
    </div>
    <div id="cmpOut"></div>
  </div>

  </div><!-- /growth -->

  <div class="page" data-p="tools">
  <div class="sec"><span class="em">🧰</span> Tools &amp; Impact (scan ke alawa)</div>
  <div class="card"><h3><span class="em">🌍</span> Impact — ApneScan ne duniya me kya kiya</h3>
    <div class="kpis" id="impact"></div>
  </div>
  <div class="grid">
    <div class="card"><h3><span class="em">🧰</span> Feature usage (kitni baar)</h3><div id="featC"></div></div>
    <div class="card"><h3><span class="em">📥</span> Feature adoption (% users)</h3><div id="featA"></div></div>
    <div class="card"><h3><span class="em">📤</span> Share method (WhatsApp / Email / Print)</h3><canvas id="sharePie" height="150"></canvas></div>
    <div class="card"><h3><span class="em">🧩</span> Multi-feature users (kitne tools use karte)</h3><canvas id="mfc" height="150"></canvas></div>
    <div class="card"><h3><span class="em">🌐</span> Bhasha pasand (Hindi / English)</h3><div id="langP"></div></div>
    <div class="card"><h3><span class="em">🎛</span> Mode pasand (Simple / Full)</h3><div id="modeP"></div></div>
  </div>

  </div><!-- /tools -->

  <div class="page" data-p="users">
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

  </div><!-- /users -->

  <div class="page" data-p="devices">
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

  </div><!-- /devices -->

  <div class="page" data-p="system">
  <div class="sec"><span class="em">💬</span> Quality, Feedback &amp; System</div>
  <div class="grid">
    <div class="card"><h3><span class="em">💬</span> Feedback (⭐ <span id="arate"></span>)</h3><div id="fb"></div></div>
    <div class="card"><h3><span class="em">💥</span> Crash reports</h3><div id="cr"></div></div>
    <div class="card"><h3><span class="em">🏅</span> Records</h3><div id="rc"></div></div>
    <div class="card"><h3><span class="em">🔒</span> Admin logins (IP)</h3><div id="al"></div></div>
  </div>

  <!-- public tools -->
  <div class="card no-print"><h3><span class="em">🌐</span> Public tools (website ke liye)</h3>
    <div style="font-size:11px;color:var(--mut);margin-bottom:8px">Ye login-free hain — sirf safe ginti dikhate hain (koi personal data nahi).</div>
    <div style="margin-bottom:8px">🔗 <b>Public stats page:</b> <a href="?public=1" target="_blank" style="color:var(--accent)">status.apnesoft.com/stats.php?public=1</a></div>
    <div style="margin-bottom:6px">📌 <b>Website widget</b> — apni site me ye code paste karo (live "X scans worldwide" dikhega):</div>
    <textarea readonly onclick="this.select()" style="width:100%;height:52px;font-family:monospace;font-size:10px">&lt;iframe src="https://status.apnesoft.com/stats.php?widget=1" style="border:0;width:420px;height:70px" scrolling="no"&gt;&lt;/iframe&gt;</textarea>
    <div style="font-size:10px;color:var(--mut);margin-top:4px">Sirf ginti chahiye (apna design)? → <a href="?widget=count" target="_blank" style="color:var(--accent)">?widget=count</a> (JSON)</div>
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

  </div><!-- /system -->

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
  Chart.defaults.font.size=10;
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
function kpi(ic,n,l,cls){return '<div class="kpi '+(cls||'')+'"><div class="ic">'+ic+'</div><div class="tx"><div class="n">'+n+'</div><div class="l">'+l+'</div></div></div>';}
document.getElementById('kpis').innerHTML=
  kpi('📄',fmt(D.total),'Total scans')+kpi('📅',fmt(D.today),'Today')+kpi('🕐',fmt(D.yesterday),'Yesterday')+
  kpi('🗓️',fmt(D.weekTotal),'This week')+kpi('📆',fmt(D.monthTotal),'This month')+kpi('📊',D.dailyAvg,'Daily avg','y')+
  kpi(D.momPct>=0?'📈':'📉',(D.momPct>=0?'+':'')+D.momPct+'%','vs last month',D.momPct>=0?'g':'r')+
  kpi('🟢',fmt(D.online),'Online now','g')+kpi('👥',fmt(D.users),'Total users','p')+kpi('✨',fmt(D.newToday),'New today','o')+
  kpi('⚡',fmt(D.dau),'Active today')+kpi('📶',fmt(D.wau),'Active 7d')+kpi('🔆',fmt(D.mau),'Active 30d')+
  kpi('🔥',fmt(D.powerUsers),'Power users','p')+kpi('🔸',fmt(D.oneTime),'One-time','r')+
  kpi('📥',fmt(D.imports),'Imports','o')+kpi('🖨️',fmt(D.prints),'Prints')+kpi('🌿',fmt(D.total),'Paper saved','g');

// AI summary
try{ if(D.aiSummary) document.getElementById('aisum').innerHTML=D.aiSummary; }catch(e){}

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

// ================= DEEP WORLDWIDE ANALYTICS (render) =================
// extra KPI row
document.getElementById('kpis2').innerHTML=
  kpi('📈',D.avgPerUser,'Avg / user','p')+
  kpi('📄',D.avgPages,'Avg pages/scan')+
  kpi(D.wowPct>=0?'📈':'📉',(D.wowPct>=0?'+':'')+D.wowPct+'%','Week/week',D.wowPct>=0?'g':'r')+
  kpi('🧲',D.stickiness+'%','Stickiness (DAU/MAU)','o')+
  kpi('🔁',D.repeatRate+'%','Repeat rate','g')+
  kpi('🔥',D.globalStreak,'Day streak','r')+
  kpi('🕐',fmt(D.lastHour),'Last hour')+
  kpi('⏱',fmt(D.last24h),'Last 24h')+
  kpi('📆',fmt(D.thisYear),'This year')+
  kpi(D.userGrowthPct>=0?'🌱':'🍂',(D.userGrowthPct>=0?'+':'')+D.userGrowthPct+'%','User growth',D.userGrowthPct>=0?'g':'r')+
  kpi('😴',D.churnRate+'%','Churn rate','r')+
  kpi('📊',(D.netGrowthMo>=0?'+':'')+D.netGrowthMo,'Net users (mo)',D.netGrowthMo>=0?'g':'r');

// milestone countdown
(function(){ var el=document.getElementById('mile');
  if(!D.nextMilestone){ el.innerHTML='<div style="color:var(--mut)">🏆 Sab milestones paar!</div>'; return; }
  el.innerHTML='<div style="font-size:22px;font-weight:800">'+fmt(D.milestoneLeft)+' <span style="font-size:13px;color:var(--mut);font-weight:600">aur '+fmt(D.nextMilestone)+' tak</span></div>'+
    '<div style="height:12px;background:var(--line);border-radius:8px;overflow:hidden;margin:10px 0 4px"><div style="height:100%;width:'+D.milestonePct+'%;background:linear-gradient(90deg,'+PAL.blue+','+PAL.aqua+')"></div></div>'+
    '<div style="color:var(--mut);font-size:12px">'+D.milestonePct+'% done · abhi '+fmt(D.total)+'</div>'; })();

// forecast
document.getElementById('fore').innerHTML='<table>'+
  '<tr><td>🔮 Agle mahine (~)</td><td style="text-align:right"><b>'+fmt(D.forecastMonth)+'</b></td></tr>'+
  '<tr><td>📅 Year-end tak (~)</td><td style="text-align:right"><b>'+fmt(D.yearEndProj)+'</b></td></tr>'+
  '<tr><td>🔥 Global streak</td><td style="text-align:right"><b>'+D.globalStreak+' din</b></td></tr>'+
  '<tr><td>🆕 Naye desh (is mahine)</td><td style="text-align:right"><b>'+((D.newCountries||[]).map(function(c){return flag(c);}).join(' ')||'—')+'</b></td></tr></table>';

// live feed
(function(){ var el=document.getElementById('feed'); var r=D.recentScans||[];
  el.innerHTML=r.length?r.map(function(s){return '<div style="padding:5px 0;border-bottom:1px solid var(--line);font-size:13px">'+(s.cc?flag(s.cc):'🌍')+' <b>'+esc(s.name||'—')+'</b> ne '+s.n+' page scan kiye <span style="color:var(--mut);float:right">'+ago(s.t)+' pehle</span></div>';}).join(''):'<div style="color:var(--mut)">— abhi koi scan nahi —</div>'; })();

if(window.Chart){
  var bc2=Chart.defaults.borderColor;
  var COb={plugins:{legend:{display:false}},scales:{x:{grid:{display:false}},y:{beginAtZero:true,grid:{color:bc2},ticks:{precision:0}}}};
  mkChart(dowc,{type:'bar',data:{labels:['Mon','Tue','Wed','Thu','Fri','Sat','Sun'],datasets:[{data:D.dow,backgroundColor:D.dow.map(function(_,i){return i>=5?PAL.orange:PAL.blue;})}]},options:COb});
  mkChart(mbmc,{type:'bar',data:{labels:D.monthByMonth.map(function(x){return x[0];}),datasets:[{data:D.monthByMonth.map(function(x){return x[1];}),backgroundColor:PAL.aqua}]},options:COb});
  var sg=D.segments||{};
  mkChart(segc,{type:'doughnut',data:{labels:['New','Regular','Power','Asleep'],datasets:[{data:[sg['new'],sg.regular,sg.power,sg.asleep],backgroundColor:[PAL.aqua,PAL.blue,PAL.violet,'#94a3b8'],borderWidth:2,borderColor:'var(--card)'}]},options:{plugins:{legend:{position:'right',labels:{boxWidth:12}}}}});
  var db=D.scanDist||{}; var dl=Object.keys(db);
  mkChart(distc,{type:'bar',data:{labels:dl,datasets:[{data:dl.map(function(k){return db[k];}),backgroundColor:PAL.blue}]},options:COb});
  mkChart(retcv,{type:'line',data:{labels:D.retCurve.map(function(x){return 'D+'+x[0];}),datasets:[{data:D.retCurve.map(function(x){return x[1];}),borderColor:PAL.aqua,backgroundColor:function(c){return grad(c,PAL.aqua);},borderWidth:2,fill:true,tension:.3,pointRadius:3,pointBackgroundColor:PAL.aqua}]},options:{plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,max:100,grid:{color:bc2},ticks:{callback:function(v){return v+'%';}}},x:{grid:{display:false}}}}});
}

// scanner brands
bars('brands',obj2rows(D.brands),false);

// 7x24 heatmap
(function(){ var el=document.getElementById('heat7'); var H=D.heat7x24||[]; var days=['Mon','Tue','Wed','Thu','Fri','Sat','Sun'];
  var mx=1; H.forEach(function(row){row.forEach(function(v){if(v>mx)mx=v;});});
  var h='<table style="border-collapse:separate;border-spacing:2px"><tr><td></td>';
  for(var hh=0;hh<24;hh++) h+='<td style="font-size:8px;color:var(--mut);text-align:center">'+hh+'</td>';
  h+='</tr>';
  H.forEach(function(row,di){ h+='<tr><td style="font-size:10px;color:var(--mut);padding-right:4px">'+days[di]+'</td>';
    row.forEach(function(v,hh){ var a=v/mx; var bg=a>0?'rgba(42,120,214,'+(0.12+a*0.88)+')':'var(--line)'; h+='<td title="'+days[di]+' '+hh+':00 — '+v+'" style="width:15px;height:15px;border-radius:3px;background:'+bg+'"></td>'; });
    h+='</tr>'; });
  el.innerHTML=h+'</table>'; })();

// world map (scatter by lat/lng)
var CC={IN:[22,79],US:[38,-97],GB:[54,-2],CA:[56,-106],AU:[-25,133],DE:[51,10],FR:[46,2],IT:[42,12],ES:[40,-4],BR:[-10,-55],RU:[61,105],CN:[35,104],JP:[36,138],KR:[36,128],ID:[-2,118],PK:[30,70],BD:[24,90],NG:[9,8],ZA:[-29,24],EG:[26,30],SA:[24,45],AE:[24,54],TR:[39,35],MX:[23,-102],AR:[-38,-63],PH:[13,122],VN:[16,108],TH:[15,101],MY:[4,102],SG:[1,104],NP:[28,84],LK:[7,81],KE:[0,38],GH:[8,-1],NL:[52,5],SE:[62,15],PL:[52,19],UA:[49,32],IR:[32,53],IQ:[33,44],QA:[25,51],KW:[29,47],OM:[21,57],NZ:[-42,174],IE:[53,-8],PT:[39,-8],CH:[47,8],BE:[50,4],AT:[47,14]};
if(window.Chart){
  var pts=obj2rows(D.scansByCountry).map(function(r){ var c=CC[r[0]]; if(!c)return null; return {x:c[1],y:c[0],r:Math.max(5,Math.min(26,Math.sqrt(r[1])*2)),cc:r[0],v:r[1]}; }).filter(Boolean);
  mkChart(wmap,{type:'bubble',data:{datasets:[{data:pts,backgroundColor:'rgba(42,120,214,.55)',borderColor:PAL.blue,borderWidth:1}]},
    options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:function(ctx){var p=ctx.raw;return p.cc+': '+fmt(p.v)+' scans';}}}},
    scales:{x:{min:-170,max:180,display:false},y:{min:-60,max:80,display:false}}}});
}

// country depth table
(function(){ var el=document.getElementById('coDeep'); var tot=D.total||1;
  var rows=obj2rows(D.scansByCountry);
  if(!rows.length){ el.innerHTML='<div style="color:var(--mut);font-size:12px">— data nahi —</div>'; return; }
  var h='<table><tr><th>Desh</th><th>Users</th><th>Scans</th><th>/user</th><th>%</th></tr>';
  rows.forEach(function(r){ var cc=r[0]; var u=(D.usersByCountry||{})[cc]||0; var pu=(D.perUserByCountry||{})[cc]||0; var pct=(100*r[1]/tot).toFixed(1);
    h+='<tr><td>'+flag(cc)+' '+cc+'</td><td>'+u+'</td><td><b>'+fmt(r[1])+'</b></td><td>'+pu+'</td><td>'+pct+'%</td></tr>'; });
  el.innerHTML=h+'</table>'; })();

// most common setting + weekday/weekend
(function(){ var el=document.getElementById('common');
  function chip(l,v,c){ return '<div style="display:inline-block;background:var(--card2);border:1px solid var(--line);border-radius:12px;padding:8px 14px;margin:4px 6px 0 0"><div style="font-size:11px;color:var(--mut)">'+l+'</div><div style="font-size:16px;font-weight:800;color:'+c+'">'+v+'</div></div>'; }
  el.innerHTML=chip('DPI',(D.topDpi&&D.topDpi[0])||'—',PAL.blue)+chip('Colour',(D.topColor&&D.topColor[0])||'—',PAL.aqua)+chip('Page size',(D.topSize&&D.topSize[0])||'—',PAL.violet); })();
(function(){ var el=document.getElementById('wwe'); var wd=D.weekdayTotal||0, we=D.weekendTotal||0, t=wd+we||1;
  el.innerHTML='<table>'+
   '<tr><td style="width:120px">💼 Weekday (Mon–Fri)</td><td><div class="bar" style="width:'+Math.max(3,100*wd/t)+'%"></div></td><td style="width:60px;text-align:right"><b>'+fmt(wd)+'</b></td></tr>'+
   '<tr><td>🏖 Weekend (Sat–Sun)</td><td><div class="bar" style="width:'+Math.max(3,100*we/t)+'%;background:'+PAL.orange+'"></div></td><td style="text-align:right"><b>'+fmt(we)+'</b></td></tr></table>'; })();

// compare tool
(function(){
  var typeEl=document.getElementById('cmpType'), abEl=document.getElementById('cmpAB'), outEl=document.getElementById('cmpOut');
  function opts(list){ return list.map(function(x){return '<option value="'+esc(x)+'">'+esc(x)+'</option>';}).join(''); }
  function countries(){ return Object.keys(D.scansByCountry||{}); }
  function versions(){ return Object.keys(D.versions||{}); }
  function build(){
    var t=typeEl.value;
    if(t==='date'){
      abEl.innerHTML='<input type="date" id="ca1"> → <input type="date" id="cb1"> &nbsp; vs &nbsp; <input type="date" id="ca2"> → <input type="date" id="cb2"> <button class="btn" onclick="cmpRun()">Compare</button>';
    } else {
      var list=(t==='country')?countries():versions();
      abEl.innerHTML='<select id="cx">'+opts(list)+'</select> vs <select id="cy">'+opts(list)+'</select> <button class="btn" onclick="cmpRun()">Compare</button>';
      var cy=document.getElementById('cy'); if(cy&&cy.options.length>1)cy.selectedIndex=1;
    }
    outEl.innerHTML='';
  }
  window.cmpRun=function(){
    var t=typeEl.value, a,b,la,lb;
    if(t==='date'){
      var a1=document.getElementById('ca1').value,b1=document.getElementById('cb1').value,a2=document.getElementById('ca2').value,b2=document.getElementById('cb2').value;
      function sum(x,y){var s=0;for(var k in D.daysMap){if(k>=x&&k<=y)s+=parseInt(D.daysMap[k])||0;}return s;}
      a=sum(a1,b1); b=sum(a2,b2); la=a1+'…'+b1; lb=a2+'…'+b2;
    } else if(t==='country'){
      var x=document.getElementById('cx').value,y=document.getElementById('cy').value;
      a=(D.scansByCountry||{})[x]||0; b=(D.scansByCountry||{})[y]||0; la=flag(x)+' '+x; lb=flag(y)+' '+y;
    } else {
      var vx=document.getElementById('cx').value,vy=document.getElementById('cy').value;
      // version -> users count (versions holds user counts)
      a=(D.versions||{})[vx]||0; b=(D.versions||{})[vy]||0; la='v'+vx; lb='v'+vy;
    }
    var mx=Math.max(a,b,1);
    function row(l,v,c){ return '<tr><td style="width:130px">'+l+'</td><td><div class="bar" style="width:'+Math.max(3,100*v/mx)+'%;background:'+c+'"></div></td><td style="width:70px;text-align:right"><b>'+fmt(v)+'</b></td></tr>'; }
    var diff=b-a, pct=a?Math.round(100*diff/a):(b?100:0);
    outEl.innerHTML='<table>'+row(la,a,PAL.blue)+row(lb,b,PAL.orange)+'</table>'+
      '<div style="margin-top:8px;font-size:13px">Antar: <b style="color:'+(diff>=0?PAL.aqua:PAL.red)+'">'+(diff>=0?'+':'')+fmt(diff)+' ('+(pct>=0?'+':'')+pct+'%)</b></div>';
  };
  typeEl.addEventListener('change',build); build();
})();

// ================= TOOLS & IMPACT (render) =================
var FEATLBL={ocr:'OCR (text)',compress:'Compress',merge:'Merge',split:'Split page',sign:'Signature',stamp:'Stamp',password:'Password',watermark:'Watermark',whatsapp:'WhatsApp share',email:'Email share',print:'Print',import:'Import',phoneimport:'Phone photo',idcard:'ID-card crop',rename:'Rename'};
function flbl(k){ return FEATLBL[k]||k; }
// impact tiles
document.getElementById('impact').innerHTML=
  kpi('📑',fmt(D.impactDocs),'Documents bane')+
  kpi('📄',fmt(D.impactPaper),'Paper digitize','g')+
  kpi('💾',(D.impactDataMB>=1024?(D.impactDataMB/1024).toFixed(1)+' GB':D.impactDataMB+' MB'),'Data bachaya','p')+
  kpi('🌳',D.impactTrees,'Ped bachaye (~)','g')+
  kpi('⏱',fmt(D.impactHours)+'h','Samay bachaya','o')+
  kpi('🖨',fmt(D.impactShares),'Print bachaye (digital)','o');
// feature counts + adoption
bars('featC',obj2rows(D.features),false,flbl);
(function(){ var a=D.featAdoption||{}; var rows=Object.keys(a).map(function(k){return [k,a[k]];}).sort(function(x,y){return y[1]-x[1];});
  var el=document.getElementById('featA'); if(!rows.length){el.innerHTML='<div style="color:var(--mut);font-size:12px">— naye app version (v132) se aayega —</div>';return;}
  el.innerHTML='<table>'+rows.map(function(r){return '<tr><td style="width:120px">'+flbl(r[0])+'</td><td><div class="bar" style="width:'+Math.max(3,r[1])+'%;background:'+PAL.aqua+'"></div></td><td style="width:44px;text-align:right">'+r[1]+'%</td></tr>';}).join('')+'</table>'; })();
// language + mode prefs
function prefBars(id,obj,map){ var el=document.getElementById(id); var rows=obj2rows(obj); if(!rows.length){el.innerHTML='<div style="color:var(--mut);font-size:12px">— v132 se aayega —</div>';return;} var t=0; rows.forEach(function(r){t+=r[1];}); el.innerHTML='<table>'+rows.map(function(r){var l=map[r[0]]||r[0]; return '<tr><td style="width:110px">'+l+'</td><td><div class="bar" style="width:'+Math.max(3,100*r[1]/t)+'%"></div></td><td style="width:70px;text-align:right"><b>'+r[1]+'</b> ('+Math.round(100*r[1]/t)+'%)</td></tr>';}).join('')+'</table>'; }
prefBars('langP',D.langDist,{hi:'🇮🇳 Hindi',en:'🇬🇧 English'});
prefBars('modeP',D.modeDist,{simple:'🟢 Simple mode',full:'🔵 Full mode'});
if(window.Chart){
  var sp=D.sharePie||{};
  mkChart(document.getElementById('sharePie'),{type:'doughnut',data:{labels:['WhatsApp','Email','Print'],datasets:[{data:[sp.whatsapp,sp.email,sp.print],backgroundColor:[PAL.green,PAL.blue,PAL.orange],borderWidth:2,borderColor:'var(--card)'}]},options:{plugins:{legend:{position:'right',labels:{boxWidth:12}}}}});
  var mfo=D.multiFeature||{}; var mk=['1','2','3','4+'];
  mkChart(document.getElementById('mfc'),{type:'bar',data:{labels:mk.map(function(k){return k+' tool';}),datasets:[{data:mk.map(function(k){return mfo[k]||0;}),backgroundColor:PAL.violet}]},options:{plugins:{legend:{display:false}},scales:{x:{grid:{display:false}},y:{beginAtZero:true,ticks:{precision:0},grid:{color:Chart.defaults.borderColor}}}}});
}

// ===== SMART SUGGESTIONS (render) =====
(function(){ var el=document.getElementById('ideas'); if(!el) return; var s=D.suggestions||[];
  if(!s.length){ el.innerHTML='<div class="card" style="color:var(--mut)">—</div>'; return; }
  var col={1:{b:'#e34948',bg:'rgba(227,73,72,.07)',l:'🔴 ZAROORI'},2:{b:'#e06a38',bg:'rgba(235,104,52,.07)',l:'🟠 MADHYAM'},3:{b:'#2a78d6',bg:'rgba(42,120,214,.06)',l:'🔵 SUJHAV'}};
  el.innerHTML=s.map(function(x){ var c=col[x.p]||col[3];
    return '<div class="card" style="border-left:4px solid '+c.b+';background:'+c.bg+'">'+
      '<div style="display:flex;align-items:flex-start;gap:11px">'+
        '<div style="font-size:23px;line-height:1">'+x.ic+'</div>'+
        '<div style="flex:1;min-width:0">'+
          '<div style="font-size:9px;font-weight:800;letter-spacing:.08em;color:'+c.b+';margin-bottom:2px">'+c.l+'</div>'+
          '<div style="font-weight:700;font-size:13px">'+esc(x.t)+'</div>'+
          '<div style="margin-top:6px;color:var(--fg2);font-size:11.5px;line-height:1.55">'+esc(x.d)+'</div>'+
        '</div>'+
      '</div></div>';
  }).join(''); })();

// ===== TABS — alag-alag pages (ek hi scroll nahi) =====
(function(){
  var pages=[].slice.call(document.querySelectorAll('.page'));
  var tabs=[].slice.call(document.querySelectorAll('.tab'));
  function resizeVisibleCharts(){ try{ if(window.Chart&&Chart.getChart){ document.querySelectorAll('.page canvas').forEach(function(cv){ if(cv.offsetParent!==null){ var ch=Chart.getChart(cv); if(ch) ch.resize(); } }); } }catch(e){} }
  function show(p){
    pages.forEach(function(el){ el.style.display=(el.getAttribute('data-p')===p)?'':'none'; });
    tabs.forEach(function(b){ b.classList.toggle('active', b.getAttribute('data-p')===p); });
    try{ localStorage.setItem('anpage',p); }catch(e){}
    resizeVisibleCharts();
  }
  tabs.forEach(function(b){ b.onclick=function(){ show(b.getAttribute('data-p')); window.scrollTo(0,0); }; });
  var start='overview'; try{ var s=localStorage.getItem('anpage'); if(s) start=s; }catch(e){}
  if(!document.querySelector('.tab[data-p="'+start+'"]')) start='overview';
  show(start);
})();

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
        $_before = intval($d['total']);
        $d['total'] = intval($d['total']) + $n;
        // milestone paar hua? -> Telegram par khushkhabri
        foreach (array(100,500,1000,5000,10000,25000,50000,100000,250000,500000,1000000) as $_m) {
            if ($_before < $_m && $d['total'] >= $_m) { tg_send("🎉 <b>Milestone!</b> ApneScan ne duniya bhar me <b>".number_format($_m)."</b> scans paar kar liye!"); break; }
        }
        $d['days'][$today] = (isset($d['days'][$today])?intval($d['days'][$today]):0) + $n;
        $hk = hour_key(); $d['hours'][$hk] = (isset($d['hours'][$hk])?intval($d['hours'][$hk]):0) + $n;
        // scan settings breakdown
        if (isset($_REQUEST['dpi'])) bump($d['dpis'],  substr($_REQUEST['dpi'],0,6), $n?1:1);
        if (isset($_REQUEST['col'])) bump($d['colors'],substr($_REQUEST['col'],0,10),1);
        if (isset($_REQUEST['sz']))  bump($d['sizes'], substr($_REQUEST['sz'],0,10),1);
        if (!empty($_REQUEST['sm']))  bump($d['scanners'], substr($_REQUEST['sm'],0,40),1);
        // scan-event count (avg pages/scan ke liye) + live feed
        $d['scanEvents'] = intval(isset($d['scanEvents'])?$d['scanEvents']:0) + 1;
        if (!isset($d['recentScans'])) $d['recentScans'] = array();
        $d['recentScans'][] = array('t'=>$now, 'name'=>substr(isset($_REQUEST['u'])?$_REQUEST['u']:'',0,40),
            'cc'=>substr(isset($_REQUEST['c'])?$_REQUEST['c']:'',0,4), 'n'=>$n);
        $d['recentScans'] = array_slice($d['recentScans'], -40);
        update_peak($d, $now, $today);
    }
} else if ($action === 'ping') {
    touch_client($d, $client, $_REQUEST, 0, $now, $today);
    update_peak($d, $now, $today);
} else if ($action === 'event') {
    touch_client($d, $client, $_REQUEST, 0, $now, $today);
    if (!empty($_REQUEST['feat'])) {
        $feat = substr($_REQUEST['feat'], 0, 20);
        bump($d['features'], $feat, 1);
        // per-feature UNIQUE users (adoption ke liye)
        if (!isset($d['featUsers'])) $d['featUsers'] = array();
        if (!isset($d['featUsers'][$feat])) $d['featUsers'][$feat] = array();
        if ($client !== '' && !in_array($client, $d['featUsers'][$feat])) {
            $d['featUsers'][$feat][] = $client;
            $d['featUsers'][$feat] = array_slice($d['featUsers'][$feat], -5000);
        }
        // per-client feature set (multi-feature analysis)
        if ($client !== '' && isset($d['clients'][$client])) {
            if (!isset($d['clients'][$client]['feats'])) $d['clients'][$client]['feats'] = array();
            $d['clients'][$client]['feats'][$feat] = intval(isset($d['clients'][$client]['feats'][$feat])?$d['clients'][$client]['feats'][$feat]:0) + 1;
        }
    }
    // numeric metrics
    if (!isset($d['metrics'])) $d['metrics'] = array();
    $kb = max(0, min(50000000, intval(isset($_REQUEST['kb'])?$_REQUEST['kb']:0)));
    $pg = max(0, min(5000, intval(isset($_REQUEST['pg'])?$_REQUEST['pg']:0)));
    if ($kb) $d['metrics']['kbSaved'] = intval(isset($d['metrics']['kbSaved'])?$d['metrics']['kbSaved']:0) + $kb;
    if ($pg && !empty($_REQUEST['feat'])) { $fk = 'pg_'.substr($_REQUEST['feat'],0,16); $d['metrics'][$fk] = intval(isset($d['metrics'][$fk])?$d['metrics'][$fk]:0) + $pg; }
} else if ($action === 'crash') {
    $_cv = substr(isset($_REQUEST['v'])?$_REQUEST['v']:'',0,10);
    $_ce = substr(isset($_REQUEST['err'])?$_REQUEST['err']:'',0,200);
    $d['crashes'][] = array('t'=>$now,'v'=>$_cv,'err'=>$_ce,'client'=>substr($client,0,40));
    $d['crashes'] = array_slice($d['crashes'], -100);
    tg_send("💥 <b>Crash report</b> (v".$_cv.")\n".htmlspecialchars($_ce));
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
maybe_backup($DATA_FILE, $d);      // roz ek auto-backup (cron ke bina bhi)
save_data($DATA_FILE, $d);

header('Content-Type: application/json');
echo json_encode(compute_stats($d, $client));
