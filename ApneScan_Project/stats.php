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

// ---- JSON STORAGE LAYER (json_storage.php) -------------------------------
// Poora storage ab dedicated module me hai: flock locking, atomic verify-
// write, rotating backups (.bak..bak4), auto-recovery, size caps, error/perf
// logging, .htaccess suraksha. Module na mila to bhi app chalti rahti hai —
// neeche ke legacy load_data/save_data khud kaam sambhal lete hain.
if (is_file(__DIR__ . '/json_storage.php')) { @require_once __DIR__ . '/json_storage.php'; }
$HAS_STORAGE = function_exists('loadJson');
if ($HAS_STORAGE) { initializeStorage($DATA_FILE); }
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
        'scanEvents'=>0,'recentScans'=>array(),'metrics'=>array(),'featUsers'=>array(),
        'blockedIPs'=>array(),
        // naye admin features ke liye
        'auditLog'=>array(),      // (3) admin actions history
        'featDaily'=>array(),     // (18) feature adoption over time [date][feat]=count
        'reqHours'=>array(),      // (26) requests per hour (load)
        'sizeDaily'=>array(),     // (26) data-file size trend
        // rename analytics ka privacy-safe meta (koi asli naam nahi — sirf ginti)
        'renameMeta'=>array('sugY'=>0,'sugN'=>0,'lenSum'=>0,'lenN'=>0,'numY'=>0,'dtY'=>0,'wdSum'=>0)
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
    if (!empty($GLOBALS['HAS_STORAGE'])) {
        // naya storage layer: flock + recovery + default — phir default_data
        // se merge (nayi keys purane data me bhi mil jayein)
        return array_merge(default_data(), loadJson($file, 'default_data'));
    }
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
    if (!empty($GLOBALS['HAS_STORAGE'])) { saveJson($file, $d, $force); return; }
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
// Returns array('cc'=>'IN','region'=>'Rajasthan','city'=>'Jaipur') — jitna mile.
function geo_country($ip) {
    $out = array('cc'=>'', 'region'=>'', 'city'=>'');
    foreach (array('HTTP_CF_IPCOUNTRY','GEOIP_COUNTRY_CODE','HTTP_X_COUNTRY_CODE') as $h) {
        if (!empty($_SERVER[$h]) && strlen($_SERVER[$h])===2 && strtoupper($_SERVER[$h])!=='XX') { $out['cc']=strtoupper($_SERVER[$h]); }
    }
    if ($ip==='' || $ip==='127.0.0.1' || strpos($ip,'192.168.')===0 || strpos($ip,'10.')===0 || strpos($ip,'172.16.')===0)
        return $out;
    $url = "http://ip-api.com/json/" . urlencode($ip) . "?fields=countryCode,regionName,city";
    $s = '';
    if (function_exists('curl_init')) {
        $ch = curl_init($url);
        curl_setopt_array($ch, array(CURLOPT_RETURNTRANSFER=>true, CURLOPT_TIMEOUT=>3, CURLOPT_CONNECTTIMEOUT=>2));
        $s = @curl_exec($ch); curl_close($ch);
    } else {
        $s = @file_get_contents($url, false, stream_context_create(array('http'=>array('timeout'=>3))));
    }
    if ($s) { $j = json_decode($s, true);
        if (isset($j['countryCode']) && strlen($j['countryCode'])===2) $out['cc']=strtoupper($j['countryCode']);
        if (!empty($j['regionName'])) $out['region']=substr($j['regionName'],0,40);
        if (!empty($j['city']))       $out['city']=substr($j['city'],0,40);
    }
    return $out;
}

function touch_client(&$d, $client, $req, $n, $now, $today) {
    if ($client === '') return true;
    if (!isset($d['clients'][$client]))
        $d['clients'][$client] = array('first'=>$now,'last'=>$now,'version'=>'','country'=>'',
            'scans'=>0,'method'=>'','name'=>'','model'=>'','note'=>'','tags'=>'','blocked'=>0,
            'days'=>array(),'active'=>array());
    $c =& $d['clients'][$client];
    if (!empty($c['blocked'])) return false;              // blocked user -> ignore
    $ip = client_ip();
    // IP se block (feature 5): is IP wale sabhi ko roko
    if ($ip !== '' && !empty($d['blockedIPs']) && in_array($ip, $d['blockedIPs'])) { $c['blocked']=1; return false; }
    $c['last'] = $now;
    if (!empty($req['v']))  $c['version'] = substr($req['v'], 0, 10);
    // DESH + RAJYA + SHEHAR: IP se (asli location) — ek baar geolocate karke cache.
    // Windows-locale (req['c']) sirf fallback (wo aksar galat: 'English (US)').
    if ($ip !== '' && (!isset($c['gip']) || $c['gip'] !== $ip)) {   // har IP par SIRF EK BAAR
        $g = geo_country($ip);
        if (!empty($g['cc']))     $c['gcc']    = $g['cc'];
        if (!empty($g['region'])) $c['region'] = $g['region'];
        if (!empty($g['city']))   $c['city']   = $g['city'];
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
    // targeted message (1) + feedback reply (2) is client ke liye
    $umsg=''; $umsgId=0; $freply=''; $freplyId=0;
    if ($client!=='' && isset($d['clients'][$client])) {
        $cc=$d['clients'][$client];
        if (!empty($cc['msg']['t'])) { $umsg=$cc['msg']['t']; $umsgId=intval($cc['msg']['id']); }
        if (!empty($cc['freply']['t'])) { $freply=$cc['freply']['t']; $freplyId=intval($cc['freply']['id']); }
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
        'umsg'=>$umsg,'umsgId'=>$umsgId,'freply'=>$freply,'freplyId'=>$freplyId,
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
    if ($_GET['widget']==='count') {
        header('Content-Type: application/json');
        header('Access-Control-Allow-Origin: *');   // website ke live counters ke liye
        header('Cache-Control: public, max-age=120');
        echo json_encode($P); exit;
    }
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

    // ---- (v3) ACTIVITY FEED API: kisi bhi din ki poori event-file (JSON) ----
    // GET ?admin=1&api=feed&date=YYYY-MM-DD  → us din ke saare events + dates
    if (isset($_GET['api']) && $_GET['api'] === 'feed') {
        header('Content-Type: application/json; charset=utf-8');
        $fdt = preg_replace('/[^0-9\-]/','', isset($_GET['date'])?$_GET['date']:date('Y-m-d'));
        if ($fdt === '') $fdt = date('Y-m-d');
        $ffp = __DIR__.'/apnescan_events/events-'.$fdt.'.jsonl';
        $fout = array();
        if (is_file($ffp)) {
            foreach (array_slice(@file($ffp, FILE_IGNORE_NEW_LINES|FILE_SKIP_EMPTY_LINES) ?: array(), -3000) as $fln) {
                $fj = json_decode($fln, true); if (is_array($fj)) $fout[] = $fj;
            }
        }
        $fds = array();
        foreach (glob(__DIR__.'/apnescan_events/events-*.jsonl') ?: array() as $fg) {
            $fds[] = substr(basename($fg), 7, 10);
        }
        rsort($fds);
        echo json_encode(array('ok'=>true,'date'=>$fdt,'events'=>$fout,
            'dates'=>array_slice($fds,0,60)), JSON_UNESCAPED_UNICODE);
        exit;
    }

    // ---- (v3.3) TIMELINE API: SAB dino ke events ek saath — naya sabse
    // pehle, page-wise. GET ?admin=1&api=tl&off=0&lim=100[&u=USER][&q=TEXT]
    if (isset($_GET['api']) && $_GET['api'] === 'tl') {
        header('Content-Type: application/json; charset=utf-8');
        $off = max(0, (int)(isset($_GET['off'])?$_GET['off']:0));
        $lim = min(500, max(1, (int)(isset($_GET['lim'])?$_GET['lim']:100)));
        $fu  = isset($_GET['u']) ? trim($_GET['u']) : '';
        $fq  = isset($_GET['q']) ? strtolower(trim($_GET['q'])) : '';
        $files = glob(__DIR__.'/apnescan_events/events-*.jsonl') ?: array();
        rsort($files);                        // nayi date wali file pehle
        $files = array_slice($files, 0, 90);  // pichhle 90 din tak
        $total = 0; $out = array(); $skip = $off;
        foreach ($files as $fp) {
            $lines = @file($fp, FILE_IGNORE_NEW_LINES|FILE_SKIP_EMPTY_LINES) ?: array();
            $lines = array_reverse($lines);   // din ke andar bhi naya pehle
            foreach ($lines as $ln) {
                $j = json_decode($ln, true); if (!is_array($j)) continue;
                if ($fu !== '' && (isset($j['u'])?$j['u']:'') !== $fu) continue;
                if ($fq !== '' && strpos(strtolower((isset($j['e'])?$j['e']:'').' '
                        .(isset($j['u'])?$j['u']:'')), $fq) === false) continue;
                $total++;
                if ($skip > 0) { $skip--; continue; }
                if (count($out) < $lim) $out[] = $j;
            }
        }
        echo json_encode(array('ok'=>true,'total'=>$total,'off'=>$off,
            'lim'=>$lim,'events'=>$out), JSON_UNESCAPED_UNICODE);
        exit;
    }
    $t0=microtime(true);
    $d=load_data($DATA_FILE);

    // login log (session me ek baar) — IP ke saath
    if (empty($_SESSION['logged'])) {
        $_SESSION['logged']=true;
        $d['adminLogins'][]=array('t'=>time(),'ip'=>isset($_SERVER['REMOTE_ADDR'])?$_SERVER['REMOTE_ADDR']:'?');
        $d['adminLogins']=array_slice($d['adminLogins'],-20); save_data($DATA_FILE,$d);
    }

    // ============================================================
    //  SELF-UPDATE — nayi stats.php seedhe admin panel se upload karke
    //  panel ko update karo (Hostinger file manager kholne ki zaroorat nahi).
    //  Suraksha: sirf logged-in admin; file ki jaanch (signature + syntax);
    //  purani file ka backup (.bak) taaki galti par wapas la sakein.
    // ============================================================
    // ---- SELF-DOWNLOAD: abhi chal rahi file download karo (replace se pehle
    //      apne paas backup rakhne ke liye). Sirf logged-in admin. ----
    if (isset($_GET['dl'])) {
        $__dlmap = array(
            'panel'     => array(__FILE__,                          'stats.php'),
            'module'    => array(__DIR__.'/json_storage.php',       'json_storage.php'),
            'panelbak'  => array(__FILE__.'.bak',                   'stats.php.bak'),
            'modulebak' => array(__DIR__.'/json_storage.php.bak',   'json_storage.php.bak'),
            'data'      => array($DATA_FILE, 'stats-data-'.date('Y-m-d-His').'.json'),
        );
        $__k = (string)$_GET['dl'];
        if (isset($__dlmap[$__k]) && is_file($__dlmap[$__k][0])) {
            header('Content-Type: application/octet-stream');
            header('Content-Disposition: attachment; filename="'.$__dlmap[$__k][1].'"');
            header('Content-Length: '.filesize($__dlmap[$__k][0]));
            header('X-Content-Type-Options: nosniff');
            readfile($__dlmap[$__k][0]);
            exit;
        }
        $_SESSION['su_msg'] = '❌ Download: file nahi mili ('.htmlspecialchars(substr($__k,0,20)).')';
        header('Location: '.strtok($_SERVER['REQUEST_URI'],'?').'?admin=1'); exit;
    }

    if ($_SERVER['REQUEST_METHOD']==='POST' && isset($_POST['act']) && $_POST['act']==='selfupdate') {
        // target: 'panel' (default) = stats.php · 'module' = json_storage.php
        $target = (isset($_POST['target']) && $_POST['target']==='module') ? 'module' : 'panel';
        $self = ($target==='module') ? __DIR__.'/json_storage.php' : __FILE__;
        $minLen = ($target==='module') ? 1500 : 3000;
        $new = '';
        if (!empty($_FILES['phpfile']['tmp_name']) && is_uploaded_file($_FILES['phpfile']['tmp_name'])) {
            $new = @file_get_contents($_FILES['phpfile']['tmp_name']);
        } elseif (isset($_POST['phpcode']) && trim($_POST['phpcode'])!=='') {
            $new = (string)$_POST['phpcode'];
        }
        $new = ($new===false) ? '' : ltrim($new, "\xEF\xBB\xBF");   // strip UTF-8 BOM
        $msg = '';
        $sigOk = ($target==='module')
            ? (strpos($new,'function loadJson')!==false && strpos($new,'function saveJson')!==false)
            : (strpos($new,'function default_data')!==false && strpos($new,'compute_stats')!==false);
        if (strlen($new) < $minLen) {
            $msg = '❌ File khaali/bahut chhoti hai — update nahi ki.';
        } elseif (strncmp(ltrim($new), '<?php', 5) !== 0) {
            $msg = '❌ Ye PHP file nahi lagti (shuru me &lt;?php nahi). Update nahi ki.';
        } elseif (!$sigOk) {
            $msg = ($target==='module')
                ? '❌ Ye json_storage.php nahi lagti (loadJson/saveJson nahi mile). Update nahi ki.'
                : '❌ Ye ApneScan ki stats.php nahi lagti (zaroori hisse nahi mile). Update nahi ki.';
        } else {
            $tmp = $self.'.new';
            if (@file_put_contents($tmp, $new) === false) {
                $msg = '❌ Temp file nahi ban paayi (folder ki write-permission check karein).';
            } else {
                // syntax check — do tarah se (jo chale): (a) exec se `php -l`,
                // warna (b) pure-PHP token_get_all(TOKEN_PARSE) jo galat syntax
                // par ParseError phenkta hai. Dono me se ek se bhi galti mile to ROK.
                $lint_ok = true; $lint_out = '';
                $checked = false;
                $disabled = array_map('trim', explode(',', (string)ini_get('disable_functions')));
                if (function_exists('exec') && !in_array('exec', $disabled)) {
                    $php = (defined('PHP_BINARY') && PHP_BINARY) ? PHP_BINARY : 'php';
                    $o = array(); $rc = 0;
                    @exec(escapeshellarg($php).' -l '.escapeshellarg($tmp).' 2>&1', $o, $rc);
                    if ($rc !== 0) { $lint_ok = false; $lint_out = implode("\n", $o); }
                    $checked = true;
                }
                if ($lint_ok && defined('TOKEN_PARSE')) {   // pure-PHP fallback (no exec needed)
                    try { token_get_all($new, TOKEN_PARSE); $checked = true; }
                    catch (\ParseError $e) { $lint_ok = false; $lint_out = $e->getMessage(); }
                    catch (\Throwable $e) { /* purane PHP: chhod do */ }
                }
                if (!$lint_ok) {
                    @unlink($tmp);
                    $msg = '❌ Nayi file me PHP syntax error hai — update ROK di (panel safe hai).<br><small>'.htmlspecialchars(substr($lint_out,0,300)).'</small>';
                } else {
                    if (is_file($self)) @copy($self, $self.'.bak');  // purani ka backup
                    if (@rename($tmp, $self) || @file_put_contents($self, $new) !== false) {
                        @unlink($tmp);
                        $msg = ($target==='module') ? '✅ Storage module update ho gaya! (nayi json_storage.php lag gayi)'
                                                    : '✅ Panel update ho gaya! (nayi stats.php lag gayi)';
                    } else {
                        @unlink($tmp);
                        $msg = '❌ File replace nahi ho paayi — permission (chmod 644 / owner) check karein.';
                    }
                }
            }
        }
        $_SESSION['su_msg'] = $msg;
        header('Location: '.strtok($_SERVER['REQUEST_URI'],'?').'?admin=1'); exit;
    }
    if ($_SERVER['REQUEST_METHOD']==='POST' && isset($_POST['act']) && $_POST['act']==='selfrestore') {
        $self = __FILE__; $bak = $self.'.bak'; $msg = '';
        if (is_file($bak)) { $msg = @copy($bak, $self) ? '↩️ Pichhla version wapas aa gaya.' : '❌ Restore fail (permission?).'; }
        else { $msg = '❌ Koi backup (.bak) nahi mila.'; }
        $_SESSION['su_msg'] = $msg;
        header('Location: '.strtok($_SERVER['REQUEST_URI'],'?').'?admin=1'); exit;
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
        // (v2) crash resolve/reopen — group-signature (md5) par fix-status
        if ($act==='fixcrash')   { if(!isset($d['crashFixed'])||!is_array($d['crashFixed']))$d['crashFixed']=array();
            $h=substr(preg_replace('/[^0-9a-f]/','',isset($_POST['sig'])?$_POST['sig']:''),0,32);
            if($h!=='')$d['crashFixed'][$h]=time(); $d['crashFixed']=array_slice($d['crashFixed'],-100,null,true); }
        if ($act==='unfixcrash') { $h=substr(preg_replace('/[^0-9a-f]/','',isset($_POST['sig'])?$_POST['sig']:''),0,32);
            if(isset($d['crashFixed'][$h]))unset($d['crashFixed'][$h]); }
        if ($act==='clearfeedback') $d['feedback']=array();
        // targeted message ek user ko (1) — app broadcast ki tarah dikhata hai
        if ($act==='umsg' && isset($d['clients'][$id])) $d['clients'][$id]['msg']=array('t'=>substr(trim($_POST['msg']),0,200),'id'=>time());
        // feedback reply (2) — feedback index par jawab
        if ($act==='freply') { $fi=intval($_POST['fi']); $rt=substr(trim($_POST['reply']),0,300);
            if (isset($d['feedback'][$fi])) { $d['feedback'][$fi]['reply']=$rt; $d['feedback'][$fi]['replyId']=time();
                $fc=isset($d['feedback'][$fi]['client'])?$d['feedback'][$fi]['client']:''; // reply us user ki app me
                if ($fc==='' && isset($d['feedback'][$fi]['name'])) { foreach($d['clients'] as $cid=>$cc){ if(trim($cc['name'])===trim($d['feedback'][$fi]['name'])){ $fc=$cid; break; } } }
                if ($fc!=='' && isset($d['clients'][$fc])) $d['clients'][$fc]['freply']=array('t'=>$rt,'id'=>time()); } }
        // bulk action (3) — kai users par ek saath
        if ($act==='bulk') { $ids=explode(',', isset($_POST['ids'])?$_POST['ids']:''); $ba=isset($_POST['ba'])?$_POST['ba']:'';
            foreach($ids as $bid){ $bid=trim($bid); if($bid==='' || !isset($d['clients'][$bid]))continue;
                if($ba==='block') $d['clients'][$bid]['blocked']=1; elseif($ba==='unblock') $d['clients'][$bid]['blocked']=0;
                elseif($ba==='tag') $d['clients'][$bid]['tags']=substr(trim($_POST['tags']),0,60);
                elseif($ba==='msg') $d['clients'][$bid]['msg']=array('t'=>substr(trim($_POST['msg']),0,200),'id'=>time()); } }
        // IP se block (5)
        if ($act==='blockip') { $bip=trim($_POST['ip']); if(filter_var($bip,FILTER_VALIDATE_IP)){ if(!isset($d['blockedIPs']))$d['blockedIPs']=array(); if(!in_array($bip,$d['blockedIPs']))$d['blockedIPs'][]=$bip;
            foreach($d['clients'] as $cid=>&$cc){ if(isset($cc['gip'])&&$cc['gip']===$bip)$cc['blocked']=1; } unset($cc); } }
        if ($act==='unblockip') { $bip=trim($_POST['ip']); if(!empty($d['blockedIPs'])) $d['blockedIPs']=array_values(array_diff($d['blockedIPs'],array($bip))); }
        // remote-config editors (rconfig me): FAQ (6), donate link (9), tips (8), flags (7)
        if (in_array($act,array('faq','donate','tips','flags'))) { if(!isset($d['rconfig'])||!is_array($d['rconfig']))$d['rconfig']=array();
            if ($act==='faq')    { $j=json_decode(isset($_POST['json'])?$_POST['json']:'',true); if(is_array($j)) $d['rconfig']['faq']=$j; }
            if ($act==='donate') { $d['rconfig']['donate_url']=substr(trim($_POST['url']),0,200); }
            if ($act==='tips')   { $lines=array_filter(array_map('trim',explode("\n", isset($_POST['tips'])?$_POST['tips']:''))); $d['rconfig']['tips']=array_slice(array_values($lines),0,20); }
            if ($act==='flags')  { $j=json_decode(isset($_POST['json'])?$_POST['json']:'',true); if(is_array($j)) $d['rconfig']['flags']=$j; } }
        // (22) feedback ticket status: open <-> resolved
        if ($act==='fstatus') { $fi=intval($_POST['fi']); if(isset($d['feedback'][$fi])){ $cur=isset($d['feedback'][$fi]['status'])?$d['feedback'][$fi]['status']:'open'; $d['feedback'][$fi]['status']=($cur==='resolved')?'open':'resolved'; } }
        // (23,24,25) rconfig content editors: canned replies, changelog, news feed
        if (in_array($act,array('canned','changelog','news'))) { if(!isset($d['rconfig'])||!is_array($d['rconfig']))$d['rconfig']=array();
            if ($act==='canned')    { $lines=array_filter(array_map('trim',explode("\n", isset($_POST['canned'])?$_POST['canned']:''))); $d['rconfig']['canned']=array_slice(array_values($lines),0,20); }
            if ($act==='changelog') { $d['rconfig']['changelog']=substr(trim(isset($_POST['changelog'])?$_POST['changelog']:''),0,4000); }
            if ($act==='news')      { $lines=array_filter(array_map('trim',explode("\n", isset($_POST['news'])?$_POST['news']:''))); $d['rconfig']['news']=array_slice(array_values($lines),0,20); } }
        // (3) audit log — har admin action ka record (kab/kaunse IP se/kya)
        if(!isset($d['auditLog'])||!is_array($d['auditLog'])) $d['auditLog']=array();
        $__adet=$id; if($act==='bulk')$__adet=(isset($_POST['ba'])?$_POST['ba']:'').' ×'.count(array_filter(array_map('trim',explode(',',isset($_POST['ids'])?$_POST['ids']:''))));
        elseif(in_array($act,array('broadcast','umsg','freply','news','canned')))$__adet=substr(trim(isset($_POST['msg'])?$_POST['msg']:(isset($_POST['reply'])?$_POST['reply']:'')),0,40);
        elseif(in_array($act,array('blockip','unblockip')))$__adet=isset($_POST['ip'])?$_POST['ip']:'';
        $d['auditLog'][]=array('t'=>time(),'ip'=>isset($_SERVER['REMOTE_ADDR'])?$_SERVER['REMOTE_ADDR']:'?','act'=>$act,'det'=>substr((string)$__adet,0,50));
        $d['auditLog']=array_slice($d['auditLog'],-200);
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
        // (v2) aur reports: monthly / country-wise / crashes / scans / feedback
        if ($e==='months'){ header('Content-Type: text/csv; charset=utf-8'); header('Content-Disposition: attachment; filename="apnescan-monthly.csv"');
            $mm=array(); foreach($d['days'] as $dt=>$c){ $k=substr($dt,0,7); if(!isset($mm[$k]))$mm[$k]=0; $mm[$k]+=intval($c); } ksort($mm);
            echo "month,scans\n"; foreach($mm as $k=>$c) echo $k.",".$c."\n"; exit; }
        if ($e==='countries'){ header('Content-Type: text/csv; charset=utf-8'); header('Content-Disposition: attachment; filename="apnescan-countries.csv"');
            $cc=array(); foreach($d['clients'] as $c){ $k=trim(isset($c['country'])?$c['country']:''); if($k==='')$k='?';
                if(!isset($cc[$k]))$cc[$k]=array(0,0); $cc[$k][0]++; $cc[$k][1]+=intval(isset($c['scans'])?$c['scans']:0); } arsort($cc);
            echo "country,users,scans\n"; foreach($cc as $k=>$v) echo '"'.str_replace('"','""',$k).'",'.$v[0].','.$v[1]."\n"; exit; }
        if ($e==='crashes'){ header('Content-Type: text/csv; charset=utf-8'); header('Content-Disposition: attachment; filename="apnescan-crashes.csv"');
            echo "time,version,error\n"; foreach((isset($d['crashes'])?$d['crashes']:array()) as $c){ echo date('Y-m-d H:i',intval(isset($c['t'])?$c['t']:0)).','.trim(isset($c['v'])?$c['v']:'').',"'.str_replace('"','""',substr(isset($c['err'])?$c['err']:'',0,300)).'"'."\n"; } exit; }
        if ($e==='scans'){ header('Content-Type: text/csv; charset=utf-8'); header('Content-Disposition: attachment; filename="apnescan-scans.csv"');
            echo "time,user,pages,scanner,country\n"; foreach((isset($d['recentScans'])?$d['recentScans']:array()) as $r){ echo date('Y-m-d H:i',intval(isset($r['t'])?$r['t']:0)).',"'.str_replace('"','""',trim(isset($r['name'])?$r['name']:'')).'",'.intval(isset($r['n'])?$r['n']:0).',"'.str_replace('"','""',trim(isset($r['sm'])?$r['sm']:'')).'",'.trim(isset($r['cc'])?$r['cc']:'')."\n"; } exit; }
        if ($e==='feedback'){ header('Content-Type: text/csv; charset=utf-8'); header('Content-Disposition: attachment; filename="apnescan-feedback.csv"');
            echo "time,name,rating,status,message\n"; foreach((isset($d['feedback'])?$d['feedback']:array()) as $f){ echo date('Y-m-d H:i',intval(isset($f['t'])?$f['t']:0)).',"'.str_replace('"','""',trim(isset($f['name'])?$f['name']:'')).'",'.intval(isset($f['stars'])?$f['stars']:0).','.(isset($f['status'])?$f['status']:'open').',"'.str_replace('"','""',substr(isset($f['msg'])?$f['msg']:'',0,300)).'"'."\n"; } exit; }
    }

    // self-update ka result message (redirect ke baad ek baar dikhao)
    $SU_MSG = '';
    if (!empty($_SESSION['su_msg'])) { $SU_MSG = $_SESSION['su_msg']; unset($_SESSION['su_msg']); }

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
        // (v2.2) feature-counts (top 8) + timeline (aakhri 10) — profile/champions ke liye
        $fc=isset($c['feats'])&&is_array($c['feats'])?$c['feats']:array(); arsort($fc);
        $fcp=array(); $fi2=0; foreach($fc as $fk=>$fv){ if($fi2++>=8)break; $fcp[]=array((string)$fk,intval($fv)); }
        $evu=array(); foreach(array_slice(isset($c['ev'])&&is_array($c['ev'])?$c['ev']:array(),-10) as $e2){ if(is_array($e2)&&count($e2)>=2)$evu[]=array(intval($e2[0]),(string)$e2[1]); }
        $userList[]=array('id'=>(string)$id,'name'=>($nm!==''?$nm:'—'),'scans'=>intval(isset($c['scans'])?$c['scans']:0),
            'featc'=>$fcp,'ev'=>array_reverse($evu),
            'first'=>intval(isset($c['first'])?$c['first']:0),'last'=>$last,'version'=>trim(isset($c['version'])?$c['version']:''),
            'country'=>trim(isset($c['country'])?$c['country']:''),'method'=>trim(isset($c['method'])?$c['method']:''),
            'model'=>trim(isset($c['model'])?$c['model']:''),'note'=>trim(isset($c['note'])?$c['note']:''),
            'tags'=>trim(isset($c['tags'])?$c['tags']:''),'blocked'=>!empty($c['blocked'])?1:0,'online'=>(($now-$last)<=300),
            'region'=>trim(isset($c['region'])?$c['region']:''),'city'=>trim(isset($c['city'])?$c['city']:''),
            'ip'=>trim(isset($c['gip'])?$c['gip']:''),
            'feats'=>array_keys(isset($c['feats'])&&is_array($c['feats'])?$c['feats']:array()),
            'active'=>array_slice(isset($c['active'])&&is_array($c['active'])?$c['active']:array(),-30),
            'daysMap'=>isset($c['days'])&&is_array($c['days'])?$c['days']:array());
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
    // asli index attach karo (fi) taaki reply/status sahi item par lage (list ulti dikhti hai)
    $fbi=array(); foreach($fb as $__k=>$__f){ $__f['fi']=$__k; $fbi[]=$__f; }
    $S['feedback']=array_reverse(array_slice($fbi,-40));
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

    // ---- Churn prediction / at-risk (13): active the, ab dheere-dheere gayab ----
    $atRisk=array();
    foreach($d['clients'] as $c){ if(!empty($c['blocked']))continue; $last=intval($c['last']); $sc=intval($c['scans']); $ad=count(isset($c['active'])?$c['active']:array()); $gap=$now-$last;
        if($sc>=3 && $ad>=2 && $gap>=4*86400 && $gap<14*86400) $atRisk[]=array('name'=>(trim($c['name'])!==''?$c['name']:'—'),'scans'=>$sc,'last'=>$last,'days'=>round($gap/86400)); }
    usort($atRisk,function($a,$b){return $b['scans']-$a['scans'];});
    $S['atRisk']=array_slice($atRisk,0,15);

    // ---- Feedback sentiment (14) ----
    $posW=array('badhiya','achha','accha','acha','best','good','great','nice','perfect','love','shukriya','thanks','superb','mast','helpful','easy','fast','sahi','अच्छा','बढ़िया');
    $negW=array('slow','bug','crash','error','problem','dikkat','kharab','kharaab','bekaar','bad','worst','hang','issue','fail','खराब','धीमा');
    $reqW=array('add','chahiye','feature','kaash','should','request','option','banao','jodo','चाहिए','बनाओ');
    $sent=array('pos'=>0,'neg'=>0,'req'=>0,'neu'=>0);
    foreach($S['feedback'] as &$f){ $m=strtolower(isset($f['msg'])?$f['msg']:''); $r=intval(isset($f['rating'])?$f['rating']:0); $s='neu'; $hp=false;$hn=false;$hr=false;
        foreach($posW as $w){ if($w!=='' && strpos($m,$w)!==false){$hp=true;break;} }
        foreach($negW as $w){ if($w!=='' && strpos($m,$w)!==false){$hn=true;break;} }
        foreach($reqW as $w){ if($w!=='' && strpos($m,$w)!==false){$hr=true;break;} }
        if($hr)$s='req'; elseif($r>=4||$hp)$s='pos'; elseif(($r>0&&$r<=2)||$hn)$s='neg'; else $s='neu';
        $f['sent']=$s; $sent[$s]++; } unset($f);
    $S['sentiment']=$sent;

    // ---- Version quality / crash-rate per version (15) ----
    $vq=array(); $vusers=$S['versions'];
    foreach($d['crashes'] as $cr){ $vv=trim(isset($cr['v'])?$cr['v']:''); if($vv!==''){ if(!isset($vq[$vv]))$vq[$vv]=array('users'=>isset($vusers[$vv])?$vusers[$vv]:0,'crashes'=>0); $vq[$vv]['crashes']++; } }
    foreach($vusers as $vv=>$u){ if(!isset($vq[$vv]))$vq[$vv]=array('users'=>$u,'crashes'=>0); }
    $vqOut=array(); foreach($vq as $vv=>$x){ $vqOut[]=array('v'=>$vv,'users'=>$x['users'],'crashes'=>$x['crashes'],'rate'=>($x['users']>0?round($x['crashes']/max(1,$x['users']),2):$x['crashes'])); }
    usort($vqOut,function($a,$b){ return strcmp($b['v'],$a['v']); });
    $S['versionQuality']=array_slice($vqOut,0,8);

    // ---- Feature correlation (16): jo X karte wo Y bhi ----
    $featCl=array();
    foreach($d['clients'] as $id=>$c){ if(!empty($c['blocked']))continue; foreach((isset($c['feats'])?$c['feats']:array()) as $ft=>$cnt){ if(!isset($featCl[$ft]))$featCl[$ft]=array(); $featCl[$ft][$id]=1; } }
    $tf=array(); foreach($featCl as $ft=>$set){ $tf[$ft]=count($set); } arsort($tf); $tf=array_slice(array_keys($tf),0,6);
    $corr=array();
    foreach($tf as $a){ foreach($tf as $b){ if(strcmp($a,$b)>=0)continue; $ca=$featCl[$a]; $cb=$featCl[$b]; $both=0; foreach($ca as $id=>$_){ if(isset($cb[$id]))$both++; }
        if(count($ca)>0 && $both>0) $corr[]=array('a'=>$a,'b'=>$b,'both'=>$both,'pct'=>round(100*$both/count($ca))); } }
    usort($corr,function($x,$y){return $y['pct']-$x['pct'];});
    $S['featCorr']=array_slice($corr,0,8);

    // ---- (18) feature adoption over time: top features ki daily series ----
    $fd=isset($d['featDaily'])&&is_array($d['featDaily'])?$d['featDaily']:array();
    ksort($fd); $fdDays=array_slice(array_keys($fd),-30);
    $ftot=array(); foreach($fd as $day=>$fl){ foreach($fl as $ft=>$cn){ $ftot[$ft]=(isset($ftot[$ft])?$ftot[$ft]:0)+$cn; } }
    arsort($ftot); $fdTop=array_slice(array_keys($ftot),0,5);
    $fdSeries=array(); foreach($fdTop as $ft){ $row=array(); foreach($fdDays as $day){ $row[]=isset($fd[$day][$ft])?intval($fd[$day][$ft]):0; } $fdSeries[$ft]=$row; }
    $S['featAdopt']=array('days'=>$fdDays,'series'=>$fdSeries);

    // ---- (15) retention heatmap: cohort (pehla din ka hafta) × hafte-baad wapsi ----
    $wk=function($ts){ return intval(floor($ts/604800)); };  // hafta-number (epoch/7d)
    $coh=array();  // cohortWeek => array(weekOffset=>uniqueUsers)
    $cohSize=array();
    foreach($d['clients'] as $c){ if(!empty($c['blocked']))continue; $first=intval(isset($c['first'])?$c['first']:0); if(!$first)continue; $cw=$wk($first);
        if(!isset($cohSize[$cw]))$cohSize[$cw]=0; $cohSize[$cw]++;
        $seen=array(); $act=isset($c['active'])&&is_array($c['active'])?$c['active']:array();
        foreach($act as $ad){ $ats=strtotime($ad); if($ats===false)continue; $off=$wk($ats)-$cw; if($off<0||$off>7)continue; $seen[$off]=1; }
        if(!isset($coh[$cw]))$coh[$cw]=array(); foreach($seen as $off=>$_){ $coh[$cw][$off]=(isset($coh[$cw][$off])?$coh[$cw][$off]:0)+1; } }
    krsort($coh); $cohOut=array(); $ci=0;
    foreach($coh as $cw=>$offs){ if($ci++>=8)break; $sz=isset($cohSize[$cw])?$cohSize[$cw]:0; if($sz<=0)continue;
        $rowp=array(); for($o=0;$o<=7;$o++){ $u=isset($offs[$o])?$offs[$o]:0; $rowp[]=($o===0)?100:round(100*$u/$sz); }
        $cohOut[]=array('label'=>date('d M',$cw*604800),'size'=>$sz,'pct'=>$rowp); }
    $S['retention']=array_reverse($cohOut);

    // ---- (27) crash grouping: same error ko group karke count ----
    $cg=array();
    $cfix=isset($d['crashFixed'])&&is_array($d['crashFixed'])?$d['crashFixed']:array();
    foreach($d['crashes'] as $cr){ $e=isset($cr['err'])?$cr['err']:''; $sig=preg_replace('/0x[0-9a-fA-F]+|\d+/','#',$e); $sig=trim(substr($sig,0,80)); if($sig==='')$sig='(unknown)';
        if(!isset($cg[$sig]))$cg[$sig]=array('sig'=>$sig,'h'=>md5($sig),'sample'=>substr($e,0,80),'count'=>0,'last'=>0,'vers'=>array());
        $cg[$sig]['count']++; $cg[$sig]['last']=max($cg[$sig]['last'],intval(isset($cr['t'])?$cr['t']:0));
        $vv=trim(isset($cr['v'])?$cr['v']:''); if($vv!=='')$cg[$sig]['vers'][$vv]=1; }
    $cgOut=array_values($cg); foreach($cgOut as &$g){ $g['vers']=implode(',',array_keys($g['vers'])); $g['fixed']=isset($cfix[$g['h']])?1:0; } unset($g);
    usort($cgOut,function($a,$b){return $b['count']-$a['count'];});
    $S['crashGroups']=array_slice($cgOut,0,15);

    // ---- (30) referral leaderboard: kisne sabse zyada 'refer' kiya ----
    $rl=array();
    foreach($d['clients'] as $c){ if(!empty($c['blocked']))continue; $rc=intval(isset($c['feats']['refer'])?$c['feats']['refer']:0); if($rc>0) $rl[]=array('name'=>(trim($c['name'])!==''?$c['name']:'—'),'n'=>$rc,'cc'=>isset($c['country'])?$c['country']:''); }
    usort($rl,function($a,$b){return $b['n']-$a['n'];});
    $S['refLeaders']=array_slice($rl,0,12);

    // ---- Rename analytics (worldwide) ----
    $rnTotal=intval(isset($d['features']['rename'])?$d['features']['rename']:0);
    $rnUsers=isset($d['featUsers']['rename'])&&is_array($d['featUsers']['rename'])?count($d['featUsers']['rename']):0;
    $rnLeaders=array(); $rnByCo=array();
    foreach($d['clients'] as $c){ if(!empty($c['blocked']))continue; $rc=intval(isset($c['feats']['rename'])?$c['feats']['rename']:0);
        if($rc>0){ $rnLeaders[]=array('name'=>(trim($c['name'])!==''?$c['name']:'—'),'n'=>$rc,'cc'=>isset($c['country'])?$c['country']:''); $co=trim($c['country']); if($co!=='')bump($rnByCo,$co,$rc); } }
    usort($rnLeaders,function($a,$b){return $b['n']-$a['n'];});
    $fd2=isset($d['featDaily'])&&is_array($d['featDaily'])?$d['featDaily']:array(); ksort($fd2);
    $rnTrend=array(); foreach(array_slice(array_keys($fd2),-30) as $day){ $rnTrend[]=array($day,intval(isset($fd2[$day]['rename'])?$fd2[$day]['rename']:0)); }
    $rnToday=intval(isset($fd2[$today]['rename'])?$fd2[$today]['rename']:0); $rnWeek=0;$rnMonth=0;
    for($k=0;$k<30;$k++){ $dk=date('Y-m-d',$now-$k*86400); $vv=intval(isset($fd2[$dk]['rename'])?$fd2[$dk]['rename']:0); $rnMonth+=$vv; if($k<7)$rnWeek+=$vv; }
    $rm=isset($d['renameMeta'])&&is_array($d['renameMeta'])?$d['renameMeta']:array();
    $rmN=intval(isset($rm['lenN'])?$rm['lenN']:0); $rmSug=intval(isset($rm['sugY'])?$rm['sugY']:0)+intval(isset($rm['sugN'])?$rm['sugN']:0);
    $S['rename']=array('total'=>$rnTotal,'users'=>$rnUsers,
        'adopt'=>($S['users']>0?round(100*$rnUsers/$S['users']):0),
        'perUser'=>($rnUsers>0?round($rnTotal/$rnUsers,1):0),
        'today'=>$rnToday,'week'=>$rnWeek,'month'=>$rnMonth,
        'leaders'=>array_slice($rnLeaders,0,12),'byCountry'=>$rnByCo,'trend'=>$rnTrend,
        'metaN'=>$rmN,
        'pickRate'=>($rmSug>0?round(100*intval($rm['sugY'])/$rmSug):0),
        'avgLen'=>($rmN>0?round($rm['lenSum']/$rmN):0),
        'avgWords'=>($rmN>0?round($rm['wdSum']/$rmN,1):0),
        'pctNum'=>($rmN>0?round(100*intval($rm['numY'])/$rmN):0),
        'pctDate'=>($rmN>0?round(100*intval($rm['dtY'])/$rmN):0));

    // ---- (26) server health: request-load + data-size trend ----
    $rh=isset($d['reqHours'])&&is_array($d['reqHours'])?$d['reqHours']:array(); ksort($rh);
    $S['reqHours']=array_slice($rh,-48,null,true);
    $sd=isset($d['sizeDaily'])&&is_array($d['sizeDaily'])?$d['sizeDaily']:array(); ksort($sd);
    $S['sizeDaily']=array_slice($sd,-30,null,true);

    // ---- (3) audit log ----
    $S['auditLog']=array_reverse(array_slice(isset($d['auditLog'])?$d['auditLog']:array(),-60));

    // health
    $S['fileKB']=file_exists($DATA_FILE)?round(filesize($DATA_FILE)/1024,1):0;
    $S['lastBackup']=intval(isset($d['lastBackup'])?$d['lastBackup']:0);
    $S['respMs']=round((microtime(true)-$t0)*1000);

    // ---- Branded PDF report (18): saaf printable page -> "Save as PDF" ----
    if (isset($_GET['report'])) {
        header('Content-Type: text/html; charset=utf-8');
        $tf=$S['features']; arsort($tf); $topf=$tf?ucfirst(key($tf)):'—';
        ?><!doctype html><html><head><meta charset="utf-8"><title>ApneScan Report — <?php echo date('d M Y'); ?></title><style>
        *{box-sizing:border-box}body{font-family:Inter,system-ui,Segoe UI,Arial;color:#111;margin:0;padding:34px;font-size:13px}
        h1{margin:0;font-size:22px;color:#12325f}.sub{color:#666;font-size:12px;margin:2px 0 20px}
        .g{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:20px}
        .k{border:1px solid #e2e8f0;border-radius:10px;padding:12px}.k .n{font-size:20px;font-weight:800;color:#1f5fb0}.k .l{color:#666;font-size:11px}
        h2{font-size:14px;margin:18px 0 8px;border-bottom:2px solid #12325f;padding-bottom:4px;color:#12325f}
        table{width:100%;border-collapse:collapse;font-size:12px}td,th{border-bottom:1px solid #eee;padding:5px 7px;text-align:left}
        .foot{margin-top:26px;color:#999;font-size:11px;text-align:center}
        @media print{body{padding:0}.noprint{display:none}}
        </style></head><body>
        <button class="noprint" onclick="window.print()" style="float:right;background:#1f5fb0;color:#fff;border:0;padding:9px 16px;border-radius:8px;font-weight:700;cursor:pointer">🖨 Save as PDF</button>
        <h1>📊 ApneScan — Worldwide Report</h1>
        <div class="sub"><?php echo date('d M Y, H:i'); ?> · ApneSoftware.com</div>
        <div class="g">
          <div class="k"><div class="n"><?php echo number_format($S['total']); ?></div><div class="l">Total scans</div></div>
          <div class="k"><div class="n"><?php echo number_format($S['today']); ?></div><div class="l">Aaj</div></div>
          <div class="k"><div class="n"><?php echo number_format($S['weekTotal']); ?></div><div class="l">Is hafte</div></div>
          <div class="k"><div class="n"><?php echo number_format($S['monthTotal']); ?></div><div class="l">Is mahine</div></div>
          <div class="k"><div class="n"><?php echo number_format($S['users']); ?></div><div class="l">Total users</div></div>
          <div class="k"><div class="n"><?php echo $S['newToday']; ?></div><div class="l">Aaj naye users</div></div>
          <div class="k"><div class="n"><?php echo $S['dailyAvg']; ?></div><div class="l">Daily avg</div></div>
          <div class="k"><div class="n"><?php echo $S['dau'].'/'.$S['mau']; ?></div><div class="l">DAU / MAU</div></div>
        </div>
        <h2>Top desh (scans)</h2><table><tr><th>Desh</th><th>Scans</th></tr>
        <?php $sc=$S['scansByCountry']; arsort($sc); $i=0; foreach($sc as $co=>$v){ if($i++>=8)break; echo "<tr><td>$co</td><td>".number_format($v)."</td></tr>"; } ?></table>
        <h2>Feature usage</h2><table><tr><th>Feature</th><th>Baar</th></tr>
        <?php $ff=$S['features']; arsort($ff); $i=0; foreach($ff as $k=>$v){ if($i++>=10)break; echo "<tr><td>".ucfirst($k)."</td><td>$v</td></tr>"; } if(!$ff)echo "<tr><td>—</td><td>—</td></tr>"; ?></table>
        <h2>App versions</h2><table><tr><th>Version</th><th>Users</th></tr>
        <?php $vv=$S['versions']; arsort($vv); foreach($vv as $k=>$v){ echo "<tr><td>v$k</td><td>$v</td></tr>"; } ?></table>
        <div class="foot">Auto-generated · sirf ginti (koi personal/document data nahi) · apnescan.apnesoft.com</div>
        <script>setTimeout(function(){window.print();},400);</script>
        </body></html><?php exit;
    }

    // admin-only extras (app API me nahi jaate — sirf is page ka $J):
    $S['recentScans300']=array_reverse(array_slice(isset($d['recentScans'])?$d['recentScans']:array(),-300));
    $S['recentEvents']=array_reverse(array_slice(isset($d['recentEvents'])?$d['recentEvents']:array(),-200));
    if (!empty($GLOBALS['HAS_STORAGE'])) { $S['health']=storageHealth($DATA_FILE); }
    $S['adminLoginsFull']=array_reverse(array_slice(isset($d['adminLogins'])?$d['adminLogins']:array(),-20));
    // ---- (v2) REAL server metrics (sirf admin-blob me; kabhi fake nahi —
    // jo function host par uplabdh nahi, wo field aati hi nahi) ----
    $sys=array('php'=>PHP_VERSION);
    if (function_exists('sys_getloadavg')) { $la=@sys_getloadavg(); if(is_array($la)&&isset($la[0])) { $sys['load']=round($la[0],2);
        // Shared host par load ka matlab CORES jaane bina nahi banta —
        // /proc/cpuinfo mile to per-core % nikaalo, warna neutral dikhao.
        $ci=@file_get_contents('/proc/cpuinfo');
        if ($ci) { $nc=preg_match_all('/^processor\s*:/m',$ci,$mm2); if($nc>0){ $sys['cores']=$nc; $sys['loadPct']=round($la[0]/$nc*100); } } } }
    $df=@disk_free_space(__DIR__); $dt=@disk_total_space(__DIR__);
    if ($df!==false && $dt) { $sys['diskFreeGB']=round($df/1073741824,2); $sys['diskTotalGB']=round($dt/1073741824,2); $sys['diskFreePct']=round($df/$dt*100,1); }
    $sys['memMB']=round(memory_get_peak_usage(true)/1048576,1);
    $sys['dataKB']=round((@filesize($DATA_FILE)?:0)/1024,1);
    $bk=array();
    foreach(array('.bak','.bak1','.bak2','.bak3','.bak4') as $sfx){ $f=$DATA_FILE.$sfx; if(@is_file($f)) $bk[]=array('n'=>basename($f),'kb'=>round((@filesize($f)?:0)/1024,1),'t'=>@filemtime($f)?:0); }
    $bdir=__DIR__.'/backups';
    if (@is_dir($bdir)) { $bl=@scandir($bdir); if(is_array($bl)){ foreach($bl as $f){ if(strpos($f,'stats-')===0) $bk[]=array('n'=>'backups/'.$f,'kb'=>round((@filesize($bdir.'/'.$f)?:0)/1024,1),'t'=>@filemtime($bdir.'/'.$f)?:0); } } }
    usort($bk,function($a,$b){return $b['t']-$a['t'];});
    $sys['backups']=array_slice($bk,0,10);
    $lg=__DIR__.'/logs/error.log'; $sys['logTail']=array();
    if (@is_file($lg)) { $lsz=@filesize($lg)?:0; $sys['logKB']=round($lsz/1024,1);
        $fh=@fopen($lg,'r'); if($fh){ if($lsz>16384) @fseek($fh,-16384,SEEK_END); $txt=@stream_get_contents($fh); @fclose($fh);
            $lines=preg_split('/\r?\n/',trim((string)$txt)); $sys['logTail']=array_slice($lines,-40); } }
    $S['sys']=$sys;
    // security: haal ki galat-password koshishen (admin-only)
    $flr=array(); foreach(array_reverse(array_slice(isset($d['failLog'])?$d['failLog']:array(),-15)) as $f){ if(is_array($f))$flr[]=array('t'=>intval(isset($f['t'])?$f['t']:0),'ip'=>isset($f['ip'])?$f['ip']:'?'); }
    $S['failLogRecent']=$flr;
    $J=json_encode($S);
    ?><!doctype html>
<html lang="hi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ApneScan — Admin Stats</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  /* ============================================================
     ApneScan Admin — Fluent (Windows-11) design system
     Dark-first enterprise palette · sidebar shell · glass topbar
     ============================================================ */
  :root{
    --bg:#0F172A;--side:#111827;--card:#1E293B;--card2:#16223A;--fg:#F8FAFC;--fg2:#CBD5E1;--mut:#94A3B8;--line:#2b3a54;
    --accent:#2563EB;--accent2:#3B82F6;--ok:#22C55E;--warn:#F59E0B;--bad:#EF4444;--purple:#8B5CF6;
    --sh:0 1px 2px rgba(0,0,0,.35),0 1px 4px rgba(0,0,0,.25);
    --sh2:0 16px 38px rgba(0,0,0,.5),0 4px 10px rgba(0,0,0,.35);
    --glass:rgba(17,24,39,.72);
    --hd:linear-gradient(120deg,#1e3a8a 0%,#2563EB 55%,#7c3aed 100%);
    --radius:14px; --tr:.22s cubic-bezier(.4,0,.2,1);
  }
  html[data-th=light]{
    --bg:#EEF2F8;--side:#ffffff;--card:#ffffff;--card2:#F1F5FB;--fg:#0F172A;--fg2:#334155;--mut:#64748B;--line:#E2E8F0;
    --sh:0 1px 2px rgba(16,24,40,.07),0 1px 3px rgba(16,24,40,.06);
    --sh2:0 14px 30px rgba(16,24,40,.14),0 3px 8px rgba(16,24,40,.07);
    --glass:rgba(255,255,255,.8);
  }
  *{box-sizing:border-box}
  body{margin:0;font-family:Inter,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:var(--bg);color:var(--fg);-webkit-font-smoothing:antialiased;font-size:11.5px;line-height:1.5}
  a{color:inherit}
  ::selection{background:rgba(59,130,246,.35)}
  :focus-visible{outline:2px solid var(--accent2);outline-offset:2px;border-radius:6px}
  /* ---------- APP SHELL: sidebar + main ---------- */
  .shell{display:grid;grid-template-columns:228px 1fr;min-height:100vh}
  aside.sb{position:sticky;top:0;height:100vh;overflow-y:auto;background:var(--side);border-right:1px solid var(--line);
    display:flex;flex-direction:column;gap:2px;padding:14px 10px;z-index:30}
  aside.sb::-webkit-scrollbar{width:0}
  .sb .brand{display:flex;align-items:center;gap:10px;padding:4px 8px 14px;font-size:15px;font-weight:800;letter-spacing:-.01em}
  .sb .brand .logo{width:34px;height:34px;border-radius:10px;background:var(--hd);display:flex;align-items:center;justify-content:center;font-size:16px;box-shadow:0 4px 12px rgba(37,99,235,.45)}
  .sb .brand small{display:block;font-size:9px;font-weight:600;color:var(--mut);letter-spacing:.05em;text-transform:uppercase}
  .sb .nlab{margin:10px 8px 4px;font-size:8.5px;font-weight:800;letter-spacing:.12em;text-transform:uppercase;color:var(--mut)}
  .tab{display:flex;align-items:center;gap:10px;width:100%;text-align:left;white-space:nowrap;border:none;background:transparent;color:var(--mut);
    font-size:11.5px;font-weight:600;padding:9px 12px;border-radius:10px;cursor:pointer;transition:all var(--tr);font-family:inherit;position:relative}
  .tab:hover{background:var(--card2);color:var(--fg)}
  .tab.active{background:linear-gradient(90deg,rgba(37,99,235,.22),rgba(37,99,235,.08));color:#fff;font-weight:700}
  html[data-th=light] .tab.active{color:var(--accent)}
  .tab.active::before{content:"";position:absolute;left:0;top:20%;bottom:20%;width:3px;border-radius:3px;background:var(--accent2)}
  .tab .tic{font-size:14px;width:20px;text-align:center;flex:none}
  .sb .jump{font-size:10.5px;padding:7px 12px}
  .sb .sfoot{margin-top:auto;padding:10px 8px 2px;font-size:9.5px;color:var(--mut);border-top:1px solid var(--line)}
  .main{min-width:0}
  /* ---------- TOPBAR (glass) ---------- */
  header{position:sticky;top:0;z-index:25;display:flex;align-items:center;gap:12px;flex-wrap:wrap;
    background:var(--glass);backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);
    border-bottom:1px solid var(--line);padding:10px 18px;color:var(--fg)}
  header .ttl{font-size:14px;font-weight:800;letter-spacing:-.01em;display:flex;align-items:center;gap:8px}
  header .live{display:flex;align-items:center;gap:6px;background:rgba(34,197,94,.12);color:var(--ok);padding:5px 11px;border-radius:30px;font-size:10.5px;font-weight:700}
  header .live .dot{width:7px;height:7px;border-radius:50%;background:var(--ok);animation:pulse 2s infinite}
  @keyframes pulse{0%{box-shadow:0 0 0 0 rgba(34,197,94,.5)}70%{box-shadow:0 0 0 6px rgba(34,197,94,0)}100%{box-shadow:0 0 0 0 rgba(34,197,94,0)}}
  .gsrch{flex:1;max-width:420px;min-width:160px;position:relative}
  .gsrch input{width:100%;padding:8px 12px 8px 32px;border-radius:10px;border:1px solid var(--line);background:var(--card2);color:var(--fg);font-size:11.5px}
  .gsrch::before{content:"🔍";position:absolute;left:10px;top:50%;transform:translateY(-50%);font-size:12px;opacity:.7}
  header .toolbar{margin-left:auto;display:flex;gap:6px;align-items:center}
  .iconbtn{width:34px;height:34px;border-radius:10px;border:1px solid var(--line);background:var(--card);color:var(--fg);cursor:pointer;font-size:14px;display:inline-flex;align-items:center;justify-content:center;text-decoration:none;transition:all var(--tr);position:relative}
  .iconbtn:hover{background:var(--card2);transform:translateY(-1px);box-shadow:var(--sh)} .iconbtn:active{transform:scale(.93)}
  .iconbtn.logout{width:auto;padding:0 12px;gap:6px;font-size:11px;font-weight:700;background:var(--accent);border-color:transparent;color:#fff}
  .iconbtn.logout:hover{background:var(--accent2)}
  .prof{display:flex;align-items:center;gap:8px;background:var(--card);border:1px solid var(--line);border-radius:30px;padding:4px 12px 4px 4px;font-weight:700;font-size:11px}
  .prof .av{width:26px;height:26px;border-radius:50%;background:var(--hd);display:flex;align-items:center;justify-content:center;font-size:12px;color:#fff}
  .nbadge{position:absolute;top:-4px;right:-4px;min-width:16px;height:16px;border-radius:9px;background:var(--bad);color:#fff;font-size:9px;font-weight:800;display:flex;align-items:center;justify-content:center;padding:0 4px}
  .ndrop{position:absolute;right:0;top:42px;width:300px;max-height:380px;overflow:auto;background:var(--card);border:1px solid var(--line);border-radius:14px;box-shadow:var(--sh2);padding:8px;display:none;z-index:40}
  .ndrop.open{display:block;animation:fadeUp .18s ease}
  .ndrop .ni{display:flex;gap:9px;padding:8px 9px;border-radius:10px;font-size:10.5px;line-height:1.45}
  .ndrop .ni:hover{background:var(--card2)}
  .ndrop .ni .t{color:var(--mut);font-size:9px}
  #burger{display:none}
  /* ---------- CONTENT ---------- */
  .wrap{max-width:none;margin:14px auto;padding:0 18px}
  .page{animation:fadeUp .28s ease}
  @keyframes fadeUp{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:none}}
  .sec{display:flex;align-items:center;gap:10px;margin:22px 2px 11px;font-size:10px;font-weight:800;letter-spacing:.10em;text-transform:uppercase;color:var(--mut)}
  .sec .em{font-size:12px}
  .sec::after{content:"";flex:1;height:1px;background:linear-gradient(90deg,var(--line),transparent)}
  .sec:first-of-type{margin-top:4px}
  /* ---------- KPI cards ---------- */
  .kpis{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:11px}
  .kpi{position:relative;overflow:hidden;display:flex;align-items:center;gap:11px;background:var(--card);
    border:1px solid var(--line);border-radius:var(--radius);padding:13px;box-shadow:var(--sh);
    transition:transform var(--tr),box-shadow var(--tr),border-color var(--tr);--kc:rgba(59,130,246,.2);--ig:linear-gradient(135deg,rgba(59,130,246,.28),rgba(37,99,235,.12))}
  .kpi::before{content:"";position:absolute;top:-45%;right:-25%;width:110px;height:110px;border-radius:50%;
    background:radial-gradient(circle,var(--kc),transparent 68%);pointer-events:none}
  .kpi:hover{transform:translateY(-3px);box-shadow:var(--sh2);border-color:rgba(59,130,246,.4)}
  .kpi .ic{position:relative;z-index:1;width:40px;height:40px;flex:none;border-radius:12px;display:flex;
    align-items:center;justify-content:center;font-size:18px;background:var(--ig);box-shadow:inset 0 0 0 1px rgba(255,255,255,.07)}
  .kpi .tx{position:relative;z-index:1;min-width:0}
  .kpi .n{font-size:19px;font-weight:800;line-height:1;letter-spacing:-.02em;color:var(--accent2)}
  .kpi .l{color:var(--mut);font-size:9.5px;margin-top:4px;font-weight:600;line-height:1.25;letter-spacing:.01em}
  .kpi.g{--kc:rgba(34,197,94,.22);--ig:linear-gradient(135deg,rgba(34,197,94,.3),rgba(34,197,94,.1))} .kpi.g .n{color:var(--ok)}
  .kpi.r{--kc:rgba(239,68,68,.22);--ig:linear-gradient(135deg,rgba(239,68,68,.3),rgba(239,68,68,.1))} .kpi.r .n{color:var(--bad)}
  .kpi.p{--kc:rgba(139,92,246,.24);--ig:linear-gradient(135deg,rgba(139,92,246,.32),rgba(139,92,246,.1))} .kpi.p .n{color:var(--purple)}
  .kpi.o{--kc:rgba(249,115,22,.22);--ig:linear-gradient(135deg,rgba(249,115,22,.3),rgba(249,115,22,.1))} .kpi.o .n{color:#fb923c}
  .kpi.y{--kc:rgba(245,158,11,.24);--ig:linear-gradient(135deg,rgba(245,158,11,.32),rgba(245,158,11,.1))} .kpi.y .n{color:var(--warn)}
  /* ---------- layout grids + cards ---------- */
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:13px}
  .grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:13px}
  @media(max-width:900px){.grid,.grid3{grid-template-columns:1fr}}
  .card{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);padding:14px 15px;margin-bottom:13px;box-shadow:var(--sh);transition:box-shadow var(--tr),border-color var(--tr)}
  .card:hover{box-shadow:var(--sh2)}
  .card h3{margin:0 0 10px;font-size:11.5px;font-weight:700;display:flex;align-items:center;gap:8px;color:var(--fg)}
  .card h3 .em{font-size:14px}
  .card canvas{height:150px!important;width:100%!important}
  /* ---------- tables ---------- */
  table{width:100%;border-collapse:collapse;font-size:10.5px}
  td,th{padding:6px 8px;border-bottom:1px solid var(--line);text-align:left}
  tbody tr{transition:background .12s} tbody tr:hover{background:var(--card2)}
  tbody tr:last-child td{border-bottom:none}
  th{cursor:pointer;color:var(--mut);font-size:9px;text-transform:uppercase;letter-spacing:.06em;font-weight:800;
     position:sticky;top:0;background:var(--card);z-index:1}
  .bar{height:8px;background:linear-gradient(90deg,var(--accent),var(--accent2));border-radius:5px;min-width:5px;box-shadow:0 1px 4px rgba(37,99,235,.4)}
  /* ---------- buttons + forms ---------- */
  .rbtn{background:var(--accent);color:#fff;border:none;border-radius:9px;padding:6px 11px;cursor:pointer;font-weight:700;margin-left:5px;text-decoration:none;display:inline-block;font-size:10.5px;box-shadow:var(--sh);transition:all var(--tr)}
  .rbtn:hover{background:var(--accent2);transform:translateY(-1px)}
  .rbtn.d{background:var(--card2);color:var(--accent2);border:1px solid var(--line)}
  input,select,textarea{padding:7px 10px;border:1px solid var(--line);border-radius:9px;font-size:11.5px;background:var(--card2);color:var(--fg);font-family:inherit;outline:none;transition:border-color var(--tr),box-shadow var(--tr)}
  input:focus,select:focus,textarea:focus{border-color:var(--accent2);box-shadow:0 0 0 3px rgba(59,130,246,.22)}
  .btn{background:var(--accent);color:#fff;border:none;border-radius:9px;padding:7px 13px;cursor:pointer;font-weight:700;font-size:11px;box-shadow:var(--sh);transition:all var(--tr)}
  .btn:hover{background:var(--accent2);transform:translateY(-1px)} .btn:active{transform:scale(.96)}
  .btn.gray{background:#475569}.btn.gray:hover{background:#64748B}
  .btn.red{background:var(--bad)}.btn.red:hover{background:#f87171}
  .banner{background:linear-gradient(90deg,rgba(245,158,11,.15),rgba(245,158,11,.08));border:1px solid rgba(245,158,11,.4);color:var(--warn);border-radius:11px;padding:10px 14px;margin-bottom:13px;font-size:11.5px;font-weight:600;display:none;box-shadow:var(--sh)}
  .btns{margin-bottom:8px;display:flex;gap:2px;flex-wrap:wrap;align-items:center}
  .foot{color:var(--mut);font-size:10px;text-align:center;margin:20px 0 10px}
  .tag{display:inline-block;background:rgba(59,130,246,.16);color:var(--accent2);border-radius:20px;padding:2px 9px;font-size:9px;margin-right:3px;font-weight:700}
  .heat{display:grid;grid-template-columns:repeat(24,1fr);gap:3px}
  .heat div{height:26px;border-radius:6px;font-size:8px;color:#fff;text-align:center;line-height:26px;font-weight:600}
  /* ---------- modal ---------- */
  #umodal{backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px)}
  #umodal>div{animation:pop .22s cubic-bezier(.34,1.4,.64,1)}
  @keyframes pop{from{opacity:0;transform:scale(.94) translateY(10px)}to{opacity:1;transform:none}}
  /* ---------- FAB ---------- */
  #fabTop{position:fixed;right:18px;bottom:18px;width:44px;height:44px;border-radius:50%;border:1px solid var(--line);
    background:var(--card);color:var(--fg);font-size:16px;cursor:pointer;box-shadow:var(--sh2);opacity:0;pointer-events:none;transition:all var(--tr);z-index:50}
  #fabTop.on{opacity:1;pointer-events:auto} #fabTop:hover{transform:translateY(-3px);background:var(--accent);color:#fff}
  /* ---------- responsive sidebar ---------- */
  #sbOverlay{display:none;position:fixed;inset:0;background:rgba(2,6,23,.55);z-index:29}
  @media(max-width:1000px){
    .shell{grid-template-columns:1fr}
    aside.sb{position:fixed;left:0;top:0;bottom:0;width:238px;transform:translateX(-105%);transition:transform .25s ease;box-shadow:var(--sh2)}
    aside.sb.open{transform:none}
    #sbOverlay.on{display:block}
    #burger{display:inline-flex}
    .gsrch{max-width:none}
  }
  /* ---------- enterprise users module ---------- */
  .ufc{border:1px solid var(--line);background:var(--card2);color:var(--mut);border-radius:20px;padding:5px 11px;font-size:10px;font-weight:700;cursor:pointer;transition:all var(--tr);font-family:inherit}
  .ufc:hover{color:var(--fg);border-color:var(--accent2)}
  .ufc.on{background:var(--accent);border-color:var(--accent);color:#fff}
  .uav{width:32px;height:32px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-weight:800;font-size:11px;color:#fff;position:relative;flex:none;box-shadow:inset 0 0 0 1px rgba(255,255,255,.15)}
  .uav.big{width:64px;height:64px;font-size:22px;border-radius:18px}
  .uav .on{position:absolute;right:-1px;bottom:-1px;width:10px;height:10px;border-radius:50%;background:var(--ok);border:2px solid var(--card)}
  .uav.big .on{width:15px;height:15px}
  .stb{display:inline-flex;align-items:center;gap:4px;border-radius:20px;padding:2px 9px;font-size:9px;font-weight:800;white-space:nowrap}
  .stb.g{background:rgba(34,197,94,.15);color:var(--ok)} .stb.y{background:rgba(245,158,11,.15);color:var(--warn)}
  .stb.r{background:rgba(148,163,184,.14);color:var(--mut)} .stb.w{background:rgba(245,158,11,.18);color:var(--warn)}
  .stb.b{background:rgba(239,68,68,.16);color:var(--bad)} .stb.p{background:rgba(139,92,246,.18);color:var(--purple)}
  .hbar{width:56px;height:7px;border-radius:4px;background:var(--line);overflow:hidden;display:inline-block;vertical-align:middle}
  .hbar i{display:block;height:100%;border-radius:4px}
  .uact{display:inline-flex;gap:3px}
  .uact button{width:24px;height:24px;border-radius:7px;border:1px solid var(--line);background:var(--card2);cursor:pointer;font-size:11px;display:inline-flex;align-items:center;justify-content:center;transition:all var(--tr)}
  .uact button:hover{border-color:var(--accent2);transform:translateY(-1px)}
  .profgrid{display:grid;grid-template-columns:1fr 1fr;gap:11px;margin-top:12px}
  @media(max-width:640px){.profgrid{grid-template-columns:1fr}}
  .pcard{background:var(--card2);border:1px solid var(--line);border-radius:12px;padding:12px}
  .pcard h4{margin:0 0 8px;font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:var(--mut);display:flex;gap:6px;align-items:center}
  .pkv{display:grid;grid-template-columns:auto 1fr;gap:3px 14px;font-size:11.5px}
  .pkv span{color:var(--mut)} .pkv b{font-weight:600;word-break:break-word}
  .ach{display:inline-flex;align-items:center;gap:5px;background:linear-gradient(135deg,rgba(245,158,11,.16),rgba(245,158,11,.06));border:1px solid rgba(245,158,11,.3);color:var(--warn);border-radius:10px;padding:4px 10px;font-size:10px;font-weight:800;margin:2px}
  .ach.off{opacity:.32;filter:grayscale(1)}
  .aii{display:flex;gap:8px;padding:6px 0;font-size:11.5px;border-bottom:1px dashed var(--line)} .aii:last-child{border:none}
  .hring{width:64px;height:64px;border-radius:50%;display:grid;place-items:center;font-weight:800;font-size:13px;flex:none}
  /* ---------- skeleton shimmer ---------- */
  .skel:empty{min-height:90px;border-radius:10px;background:linear-gradient(100deg,var(--card2) 40%,rgba(148,163,184,.14) 50%,var(--card2) 60%);background-size:200% 100%;animation:skel 1.2s infinite}
  @keyframes skel{to{background-position:-200% 0}}
  /* ---------- command palette ---------- */
  #cpal{position:fixed;inset:0;background:rgba(2,6,23,.6);backdrop-filter:blur(6px);z-index:120;display:none;align-items:flex-start;justify-content:center;padding-top:12vh}
  #cpal.open{display:flex}
  #cpal .box{width:min(560px,92vw);background:var(--card);border:1px solid var(--line);border-radius:16px;box-shadow:var(--sh2);overflow:hidden;animation:pop .18s ease}
  #cpal input{width:100%;border:none;border-bottom:1px solid var(--line);border-radius:0;padding:14px 16px;font-size:14px;background:transparent}
  #cpal .res{max-height:320px;overflow:auto;padding:6px}
  #cpal .ri{display:flex;gap:10px;align-items:center;padding:9px 12px;border-radius:10px;cursor:pointer;font-size:12px}
  #cpal .ri:hover,#cpal .ri.sel{background:var(--card2)}
  #cpal .ri .k{margin-left:auto;color:var(--mut);font-size:9.5px}
  /* ---------- toasts ---------- */
  #toasts{position:fixed;right:18px;top:70px;display:flex;flex-direction:column;gap:8px;z-index:130}
  .toast{background:var(--card);border:1px solid var(--line);border-left:4px solid var(--accent2);border-radius:12px;padding:10px 14px;box-shadow:var(--sh2);font-size:11.5px;font-weight:600;animation:fadeUp .25s ease;max-width:320px}
  .toast.ok{border-left-color:var(--ok)} .toast.bad{border-left-color:var(--bad)}
  @media print{.toolbar,.btns,.no-print,aside.sb,#fabTop,.gsrch{display:none}.shell{display:block}.card{break-inside:avoid;box-shadow:none}.page{display:block!important}}
</style></head><body>
<div class="shell">
<!-- ============ SIDEBAR (Fluent) ============ -->
<aside class="sb no-print" id="sb" aria-label="Navigation">
  <div class="brand">
    <div class="logo">📊</div>
    <div>ApneScan <small>Admin Center</small></div>
  </div>
  <div class="nlab">Portal</div>
  <button class="tab" data-p="overview"><span class="tic">🏠</span> Dashboard</button>
  <button class="tab" data-p="live"><span class="tic">🟢</span> Live Monitoring</button>
  <button class="tab jump2" data-p="useract"><span class="tic">📜</span> User Activity</button>
  <button class="tab" data-p="users"><span class="tic">👥</span> Users</button>
  <button class="tab" data-p="hw"><span class="tic">🖨</span> Devices</button>
  <button class="tab" data-p="scans"><span class="tic">📄</span> Scans</button>
  <button class="tab" data-p="trends"><span class="tic">📈</span> Analytics</button>
  <div class="nlab">Engage</div>
  <button class="tab jump" onclick="jumpTo('overview','cardBroadcast')"><span class="tic">📣</span> Broadcast Center</button>
  <button class="tab jump" onclick="jumpTo('system','cardFeedback')"><span class="tic">💬</span> Feedback Center</button>
  <button class="tab jump" onclick="jumpTo('system','cardCrashes')"><span class="tic">💥</span> Crash Center</button>
  <div class="nlab">Operations</div>
  <button class="tab" data-p="reports"><span class="tic">📑</span> Reports</button>
  <button class="tab jump" onclick="jumpTo('system','cardStorage')"><span class="tic">🗄️</span> Backup Manager</button>
  <button class="tab jump" onclick="jumpTo('system','cardUpdate')"><span class="tic">⚙️</span> Settings</button>
  <button class="tab" data-p="health"><span class="tic">❤️</span> System Health</button>
  <button class="tab" data-p="logs"><span class="tic">📋</span> Activity Logs</button>
  <div class="nlab">More analytics</div>
  <button class="tab jump2" data-p="growth"><span class="tic">🔁</span> Growth</button>
  <button class="tab jump2" data-p="devices"><span class="tic">🌍</span> Activity Feed</button>
  <button class="tab jump2" data-p="tools"><span class="tic">🧰</span> Tools &amp; Impact</button>
  <button class="tab jump2" data-p="ideas"><span class="tic">💡</span> Suggestions</button>
  <button class="tab jump2" data-p="system"><span class="tic">🖥</span> System</button>
  <div class="sfoot">🟢 JSON storage · flock + backups<br>ApneSoftware.com</div>
</aside>
<div id="sbOverlay" onclick="sbToggle(false)"></div>
<div class="main">
<!-- ============ TOPBAR (glass) ============ -->
<header>
  <button class="iconbtn no-print" id="burger" onclick="sbToggle()" aria-label="Menu">☰</button>
  <div class="ttl">Dashboard</div>
  <div class="live no-print"><span class="dot"></span> Live · <span id="tm"></span> · <span id="rt" style="opacity:.75"></span></div>
  <div class="gsrch no-print"><input id="gsearch" placeholder="Search users, country, scanner… (Ctrl+K)" aria-label="Global search"></div>
  <div class="toolbar no-print">
    <button class="iconbtn" id="bellBtn" onclick="bellToggle()" title="Notifications" aria-label="Notifications">🔔<span class="nbadge" id="nbadge" style="display:none">0</span>
      <div class="ndrop" id="ndrop" onclick="event.stopPropagation()"></div>
    </button>
    <button class="iconbtn" onclick="toggleTh()" id="thbtn" title="Theme">☀️</button>
    <button class="iconbtn" onclick="location.reload()" title="Refresh">🔄</button>
    <button class="iconbtn" onclick="window.print()" title="Print">🖨</button>
    <span class="prof" title="Admin"><span class="av">👤</span> Admin</span>
    <a class="iconbtn logout" href="?logout=1" title="Logout">🔓 Logout</a>
  </div>
</header>
<div class="wrap">
  <?php if ($SU_MSG !== ''): ?>
  <div style="margin:0 0 12px;padding:11px 14px;border-radius:10px;font-size:13px;font-weight:600;<?php echo (strpos($SU_MSG,'✅')!==false||strpos($SU_MSG,'↩')!==false)?'background:#e6f7ef;color:#0a7a4f;border:1px solid #9fdcbf':'background:#fdecea;color:#b3261e;border:1px solid #f5b5ad'; ?>"><?php echo $SU_MSG; ?></div>
  <?php endif; ?>
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
    <div class="card" id="cardBroadcast"><h3><span class="em">📣</span> Broadcast — sab users ki app me message dikhao</h3>
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

  <div class="card"><h3><span class="em">🔥</span> Retention cohort heatmap <span style="color:var(--mut);font-weight:500;font-size:11px">— har hafte judне waale kitne hafte tak tike (%)</span></h3>
    <div style="overflow-x:auto"><div id="retgrid"></div></div>
  </div>

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

  <div class="card"><h3><span class="em">📈</span> Feature adoption over time <span style="color:var(--mut);font-weight:500;font-size:11px">— roz kaunsa tool kitna use hua (top 5)</span></h3>
    <canvas id="featAdopt" height="160"></canvas>
  </div>

  <div class="sec"><span class="em">✏️</span> Rename Analytics (worldwide)</div>
  <div class="kpis" id="rnKpis" style="margin-bottom:14px"></div>
  <div class="grid">
    <div class="card"><h3><span class="em">📈</span> Rename trend (30 din)</h3><canvas id="rnTrend" height="150"></canvas></div>
    <div class="card"><h3><span class="em">🏆</span> Top renamers</h3><div id="rnLeaders"></div></div>
    <div class="card"><h3><span class="em">🗺</span> Rename by country</h3><div id="rnByCo"></div></div>
    <div class="card"><h3><span class="em">🧠</span> Naam ki aadatein <span style="color:var(--mut);font-weight:500;font-size:11px">— v141 ke baad bharega</span></h3><div id="rnHabits"></div></div>
  </div>

  </div><!-- /tools -->

  <div class="page" data-p="users">
  <div class="sec"><span class="em">👤</span> Users</div>
  <div class="card"><h3><span class="em">👤</span> Saare users (<span id="ucount"></span>) <span style="color:var(--mut);font-weight:500;font-size:11px">— header pe click = sort, naam pe click = details/manage</span></h3>
    <input id="usearch" class="no-print" placeholder="🔍 naam / scanner / desh / region / version / tag se dhoondo…" style="width:100%;margin-bottom:8px">
    <div class="no-print" id="ufbar" style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:9px;align-items:center">
      <button class="ufc on" data-f="all">Sab</button>
      <button class="ufc" data-f="online">🟢 Online</button>
      <button class="ufc" data-f="offline">🔴 Offline</button>
      <button class="ufc" data-f="today">📅 Aaj active</button>
      <button class="ufc" data-f="oldver">⚠ Purana version</button>
      <button class="ufc" data-f="inactive">😴 14d+ inactive</button>
      <button class="ufc" data-f="heavy">🔥 Heavy users</button>
      <button class="ufc" data-f="attention">🚨 Dhyan chahiye</button>
      <select id="ufCountry" style="font-size:10.5px;padding:5px 8px"><option value="">🌍 Sab desh</option></select>
    </div>
    <div class="no-print" style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-bottom:8px">
      <span style="font-size:11px;color:var(--mut)">📁 Segments:</span>
      <span id="segbar" style="display:flex;flex-wrap:wrap;gap:5px"></span>
      <button class="btn gray" style="padding:3px 9px;font-size:11px" onclick="saveSeg()">➕ Abhi ka filter save karo</button>
    </div>
    <form method="post" id="bulkform" class="no-print" style="display:none;gap:6px;align-items:center;flex-wrap:wrap;margin-bottom:8px;padding:8px;background:var(--card2);border-radius:8px">
      <input type="hidden" name="act" value="bulk"><input type="hidden" name="ids" id="bulkids">
      <b style="font-size:12px"><span id="bulkn">0</span> chune</b>
      <select name="ba" id="bulkba" style="font-size:11px"><option value="tag">🏷 Tag</option><option value="msg">📨 Message</option><option value="block">🚫 Block</option><option value="unblock">✅ Unblock</option></select>
      <input name="tags" id="bulktags" placeholder="Tags: VIP,Hospital" style="font-size:11px;width:130px">
      <input name="msg" id="bulkmsg" placeholder="Message…" style="font-size:11px;width:150px;display:none">
      <button class="btn" onclick="return bulkGo()">Lagao</button>
    </form>
    <div style="overflow:auto;max-height:460px"><table id="utable"></table></div>
  </div>
  <div class="grid" style="margin-top:13px">
    <div class="card"><h3><span class="em">🏆</span> Feature champions — kaun kya SABSE ZYADA karta hai</h3><div id="uchamp" class="skel"></div></div>
    <div class="card"><h3><span class="em">🕒</span> User activity feed — kaun abhi kya kar raha hai</h3><div id="uactfeed" class="skel" style="max-height:300px;overflow:auto"></div></div>
  </div>

  
<div id="umodal" class="no-print" onclick="if(event.target===this)closeUser()" style="display:none;position:fixed;inset:0;background:rgba(15,23,42,.55);z-index:99;align-items:center;justify-content:center">
    <div style="background:var(--card);color:var(--fg);border-radius:16px;max-width:820px;width:96%;max-height:92vh;overflow:auto;padding:22px;box-shadow:0 24px 70px rgba(0,0,0,.45);position:relative">
      <button onclick="closeUser()" style="position:absolute;top:12px;right:14px;border:none;background:var(--line);border-radius:8px;width:30px;height:30px;cursor:pointer;font-size:16px">✕</button>
      <div id="umbody"></div>
    </div>
  </div>

  </div><!-- /users -->

  <div class="page" data-p="devices">
  <div class="sec"><span class="em">🧰</span> Activity, Usage &amp; Devices</div>
  <div class="card" style="margin-bottom:13px"><h3><span class="em">📜</span> Din-bhar ki POORI feed — har click <span style="color:var(--mut);font-weight:500;font-size:11px">(app v183+ · roz ki alag file · kuch nahi mitta)</span></h3>
    <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin:8px 0">
      <select id="dfDate"></select>
      <select id="dfUser"><option value="">👥 Sab users</option></select>
      <input id="dfQ" placeholder="🔍 kaam se dhoondo… (scan / save / print)" style="flex:1;min-width:140px">
      <button class="btn" onclick="dfLoad()">🔄 Load</button>
    </div>
    <div id="dfKpi" style="color:var(--mut);font-size:11px;margin-bottom:6px"></div>
    <div style="color:var(--mut);font-size:10px;margin-bottom:6px">💡 Ek user chuno to uski POORI din ki kahani sessions me bant kar dikhti hai (subah → raat)</div>
    <div id="dfList" style="max-height:420px;overflow:auto"></div>
  </div>
  <div class="grid">
    <div class="card"><h3><span class="em">🟢</span> Abhi online (<span id="oncount"></span>)</h3><div id="onlist"></div></div>
    <div class="card"><h3><span class="em">🕒</span> Recently active</h3><div id="recent"></div></div>
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

  <div class="page" data-p="useract">
  <div class="sec"><span class="em">📜</span> User Activity — kaun, kya, kab (poori kahani)</div>

  <!-- ==== DO BHAAG BARAABAR ME: beech ki patti kheench kar chhota-bada ==== -->
  <div id="uaSplit" style="display:flex;align-items:stretch;gap:0">

  <!-- BHAAG 1 (LEFT): POORI TIMELINE — sab din, naya sabse upar -->
  <div id="uaLeft" style="flex:0 0 52%;min-width:260px;overflow:hidden">
  <div class="card" style="height:100%;display:flex;flex-direction:column;box-sizing:border-box">
    <h3><span class="em">🕒</span> Poori timeline (sab din) <span id="tlInfo" style="color:var(--mut);font-weight:500;font-size:11px"></span></h3>
    <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:8px">
      <input id="tlQ" placeholder="🔍 user ya kaam se dhoondo… (poore itihaas me)" style="flex:1;min-width:140px">
      <button class="btn" onclick="tlGo(0)">🔄</button>
      <button class="btn gray" id="tlPrev" onclick="tlGo(_tl.page-1)">◀ Naye</button>
      <span id="tlPage" style="font-size:11px;color:var(--mut)">Page 1</span>
      <button class="btn gray" id="tlNext" onclick="tlGo(_tl.page+1)">Purane ▶</button>
    </div>
    <div style="color:var(--mut);font-size:10px;margin-bottom:6px">💡 Sabse naya sabse upar · 100/page · 90 din tak · beech ki patti kheench kar chaudai badlo</div>
    <div id="tlList" style="flex:1;min-height:200px;overflow:auto"></div>
  </div>
  </div>

  <!-- beech ki kheenchne wali patti -->
  <div id="uaGut" title="Kheench kar dono taraf chhota-bada karo"
       style="flex:0 0 12px;cursor:col-resize;display:flex;align-items:center;justify-content:center;user-select:none">
    <div style="width:4px;height:64px;border-radius:3px;background:rgba(148,163,184,.35)"></div>
  </div>

  <!-- BHAAG 2 (RIGHT): DIN KA SARAANSH -->
  <div id="uaRight" style="flex:1;min-width:260px;overflow:hidden">
  <div class="sec" style="font-size:13px;margin-top:2px"><span class="em">📊</span> Din ka saraansh</div>
  <div class="kpis" id="uaKpis" style="margin-bottom:13px"></div>
  <div class="card" style="margin-bottom:13px">
    <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
      <select id="uaDate"></select>
      <select id="uaUser"><option value="">👥 Sab users</option></select>
      <select id="uaType"><option value="">🧩 Sab kaam</option><option value="feat">⚙️ Features (scan/save/print…)</option><option value="btn">🔘 Toolbar buttons</option><option value="menu">📋 Menu</option><option value="nav">🧭 Sidebar/Dashboard</option><option value="app">🟢 App khuli/band</option></select>
      <input id="uaQ" placeholder="🔍 dhoondo…" style="flex:1;min-width:110px">
      <button class="btn" onclick="uaLoad()">🔄</button>
      <button class="btn gray" onclick="uaCSV()">⬇ CSV</button>
    </div>
  </div>
  <div class="card" style="margin-bottom:13px"><h3><span class="em">👥</span> Har user ka saraansh <span style="color:var(--mut);font-weight:500;font-size:11px">— naam par click = us par filter</span></h3><div id="uaUsers" style="max-height:260px;overflow:auto"></div></div>
  <div class="card" style="margin-bottom:13px"><h3><span class="em">🕐</span> Kis ghante kitna kaam</h3><div id="uaHours"></div></div>
  <div class="card"><h3><span class="em">🔥</span> Sabse zyada hue kaam</h3><div id="uaTop"></div></div>
  </div>

  </div><!-- /uaSplit -->
  </div><!-- /useract -->

  <div class="page" data-p="system">
  <div class="sec"><span class="em">💬</span> Quality, Feedback &amp; System</div>
  <div class="grid">
    <div class="card" id="cardFeedback"><h3><span class="em">💬</span> Feedback (⭐ <span id="arate"></span>) <span id="sentsum" style="font-weight:500;font-size:11px"></span></h3><div id="fb"></div></div>
    <div class="card" id="cardCrashes"><h3><span class="em">💥</span> Crash reports</h3><div id="cr"></div></div>
    <div class="card"><h3><span class="em">😟</span> Churn-risk (chhutne waale) <span style="color:var(--mut);font-weight:500;font-size:11px">— active the, 4–14 din se gayab</span></h3><div id="atr"></div></div>
    <div class="card"><h3><span class="em">🩺</span> Version quality (crash-rate)</h3><div id="vq"></div></div>
    <div class="card"><h3><span class="em">🔗</span> Feature jodi (jo X karte wo Y bhi)</h3><div id="fcorr"></div></div>
    <div class="card"><h3><span class="em">💥</span> Crash groups (ek-jaisi errors) <span style="color:var(--mut);font-weight:500;font-size:11px">— milti-julti group karke</span></h3><div id="cg"></div></div>
    <div class="card"><h3><span class="em">📣</span> Referral leaderboard <span style="color:var(--mut);font-weight:500;font-size:11px">— kisne sabse zyada share kiya</span></h3><div id="rf"></div></div>
    <div class="card"><h3><span class="em">🏆</span> Top users leaderboard <span style="color:var(--mut);font-weight:500;font-size:11px">— sabse zyada scan</span></h3><div id="lb"></div></div>
    <div class="card"><h3><span class="em">🩺</span> Server health (load + data-size)</h3><canvas id="hlc" height="120"></canvas><div id="hl" style="margin-top:6px"></div></div>
    <div class="card"><h3><span class="em">📜</span> Audit log <span style="color:var(--mut);font-weight:500;font-size:11px">— admin ke saare actions</span></h3><div style="max-height:260px;overflow:auto"><div id="au"></div></div></div>
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

  <!-- remote-config editors (app inko live padhta hai — bina naye build ke badlав) -->
  <div class="grid no-print">
    <div class="card"><h3><span class="em">❓</span> App FAQ (in-app madad)</h3>
      <div style="font-size:11px;color:var(--mut);margin-bottom:6px">JSON list — har item <code>{"q":"sawal","a":"jawab"}</code>. App ke FAQ me neeche add ho jaata hai.</div>
      <form method="post"><input type="hidden" name="act" value="faq">
        <textarea name="json" id="faqbox" rows="5" style="width:100%;font-family:monospace;font-size:10px" placeholder='[{"q":"...","a":"..."}]'></textarea>
        <button class="btn" style="margin-top:6px">FAQ save karo</button>
      </form>
    </div>
    <div class="card"><h3><span class="em">💛</span> Donate link</h3>
      <div style="font-size:11px;color:var(--mut);margin-bottom:6px">"Support / Donate" button isi URL par jaata hai (UPI/website).</div>
      <form method="post"><input type="hidden" name="act" value="donate">
        <input name="url" id="donbox" placeholder="https://... ya upi://pay?pa=..." style="width:100%">
        <button class="btn" style="margin-top:6px">Link save karo</button>
      </form>
    </div>
    <div class="card"><h3><span class="em">💡</span> App tips (rotating)</h3>
      <div style="font-size:11px;color:var(--mut);margin-bottom:6px">Har line = ek tip. App status-bar/start par ghoomti hai (max 20).</div>
      <form method="post"><input type="hidden" name="act" value="tips">
        <textarea name="tips" id="tipbox" rows="4" style="width:100%;font-size:11px" placeholder="Tip 1&#10;Tip 2"></textarea>
        <button class="btn" style="margin-top:6px">Tips save karo</button>
      </form>
    </div>
    <div class="card"><h3><span class="em">🚩</span> Feature flags</h3>
      <div style="font-size:11px;color:var(--mut);margin-bottom:6px">JSON — feature on/off. Jaise <code>{"phonescan":true,"donate":false}</code>.</div>
      <form method="post"><input type="hidden" name="act" value="flags">
        <textarea name="json" id="flagbox" rows="4" style="width:100%;font-family:monospace;font-size:10px" placeholder='{"phonescan":true}'></textarea>
        <button class="btn" style="margin-top:6px">Flags save karo</button>
      </form>
    </div>
    <div class="card"><h3><span class="em">📋</span> Ready jawab (canned replies)</h3>
      <div style="font-size:11px;color:var(--mut);margin-bottom:6px">Har line = ek ready jawab. Feedback reply me dropdown se chun sako (bar-bar na likhna pade).</div>
      <form method="post"><input type="hidden" name="act" value="canned">
        <textarea name="canned" id="cannedbox" rows="4" style="width:100%;font-size:11px" placeholder="Shukriya! Hum jald theek karenge.&#10;Ye feature agle update me aa raha hai."></textarea>
        <button class="btn" style="margin-top:6px">Replies save karo</button>
      </form>
    </div>
    <div class="card"><h3><span class="em">📰</span> Changelog (kya naya) <span style="color:var(--mut);font-weight:500;font-size:11px">— app/website me dikhega</span></h3>
      <div style="font-size:11px;color:var(--mut);margin-bottom:6px">Naye update me kya-kya aaya, yahan likho. (App-side display v141 me.)</div>
      <form method="post"><input type="hidden" name="act" value="changelog">
        <textarea name="changelog" id="chlogbox" rows="4" style="width:100%;font-size:11px" placeholder="v141: Naya phone-scan, tez OCR…"></textarea>
        <button class="btn" style="margin-top:6px">Changelog save karo</button>
      </form>
    </div>
    <div class="card"><h3><span class="em">📢</span> News feed (app me khabrein)</h3>
      <div style="font-size:11px;color:var(--mut);margin-bottom:6px">Har line = ek khabar/announcement. App me feed ki tarah chalegi. (App-side display v141 me.)</div>
      <form method="post"><input type="hidden" name="act" value="news">
        <textarea name="news" id="newsbox" rows="4" style="width:100%;font-size:11px" placeholder="Naya update aa gaya!&#10;Ab 12 bhasha me OCR."></textarea>
        <button class="btn" style="margin-top:6px">News save karo</button>
      </form>
    </div>
  </div>

  <!-- custom report builder (20) -->
  <div class="card no-print"><h3><span class="em">🧮</span> Apni report banao (custom builder)</h3>
    <div style="font-size:11px;color:var(--mut);margin-bottom:8px">Jo metrics chahiye tick karo → apni summary bane. Print/PDF bhi kar sako.</div>
    <div id="rbpick" style="display:flex;flex-wrap:wrap;gap:6px 14px;margin-bottom:10px"></div>
    <button class="btn" onclick="buildReport()">Report banao</button>
    <div id="rbout" style="margin-top:12px"></div>
  </div>

  <!-- printable report -->
  <div class="card no-print"><h3><span class="em">📄</span> Report (PDF)</h3>
    <div style="font-size:11px;color:var(--mut);margin-bottom:8px">Ek-page ka worldwide summary — KPIs, top desh, feature usage, versions. Print → "Save as PDF".</div>
    <a class="btn" href="?admin=1&report=1" target="_blank">📄 Report kholo / PDF banao</a>
  </div>

  <!-- self-update: nayi stats.php seedhe yahin se -->
  <div class="card no-print" id="cardUpdate" style="border:1px solid var(--accent)"><h3><span class="em">🔄</span> Panel update (nayi stats.php yahin se)</h3>
    <div style="font-size:11px;color:var(--mut);margin-bottom:8px">Jab bhi nayi <b>stats.php</b> mile, yahan upload kar do — panel khud update ho jaayega (Hostinger file manager kholne ki zaroorat nahi). Purani file ka backup apne aap ban jaata hai, aur syntax galat ho to update ROK di jaati hai (panel safe rehta hai).</div>
    <form method="post" enctype="multipart/form-data" onsubmit="return confirm('Nayi stats.php se panel update karein? (purani ka backup ban jaayega)')">
      <input type="hidden" name="act" value="selfupdate">
      <select name="target" style="font-size:12px"><option value="panel">stats.php (panel)</option><option value="module">json_storage.php (storage)</option></select>
      <input type="file" name="phpfile" accept=".php,text/x-php,application/x-php" style="font-size:12px">
      <button class="btn" style="margin-left:6px">⬆️ Upload &amp; update</button>
    </form>
    <details style="margin-top:8px"><summary style="cursor:pointer;font-size:11px;color:var(--mut)">…ya file ki jagah code paste karo</summary>
      <form method="post" onsubmit="return confirm('Paste ki hui code se panel update karein?')" style="margin-top:6px">
        <input type="hidden" name="act" value="selfupdate">
        <textarea name="phpcode" rows="4" style="width:100%;font-family:monospace;font-size:10px" placeholder="&lt;?php … poori stats.php ka code …"></textarea>
        <button class="btn" style="margin-top:6px">Code se update</button>
      </form>
    </details>
    <form method="post" onsubmit="return confirm('Pichhle version par wapas jaayein?')" style="margin-top:8px">
      <input type="hidden" name="act" value="selfrestore">
      <button class="btn gray">↩️ Pichhla version wapas lao (restore)</button>
    </form>
    <div style="margin-top:10px;padding-top:8px;border-top:1px dashed var(--line)">
      <div style="font-size:11px;color:var(--mut);margin-bottom:5px">⬇️ <b>Abhi ki files download karo</b> — replace karne se pehle apna backup rakh lo:</div>
      <a class="btn gray" style="text-decoration:none" href="?admin=1&amp;dl=panel">⬇ stats.php</a>
      <a class="btn gray" style="text-decoration:none" href="?admin=1&amp;dl=module">⬇ json_storage.php</a>
      <a class="btn gray" style="text-decoration:none" href="?admin=1&amp;dl=data">⬇ stats.json (data)</a>
      <a class="btn gray" style="text-decoration:none;font-size:11px" href="?admin=1&amp;dl=panelbak">stats.php.bak</a>
      <a class="btn gray" style="text-decoration:none;font-size:11px" href="?admin=1&amp;dl=modulebak">json_storage.php.bak</a>
    </div>
    <div style="font-size:10px;color:var(--mut);margin-top:8px">🔐 Suraksha: ye sirf login ke baad chalta hai. Isliye <b>admin password mazboot rakhein</b> (Settings me <code>$ADMIN_PASS</code> badlein) — warna koi aur bhi panel badal sakta hai.</div>
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

  <?php if (!empty($GLOBALS['HAS_STORAGE'])) { $SH = storageHealth($DATA_FILE); ?>
  <!-- storage health (json_storage.php module) -->
  <div class="card" id="cardStorage"><h3><span class="em">🗄️</span> Storage health (JSON layer)</h3>
    <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:8px;font-size:12px">
      <div>📄 JSON size<br><b><?php echo $SH['fileKB']; ?> KB</b></div>
      <div>📁 Backups size (.bak×5 + daily)<br><b><?php echo $SH['backupKB']; ?> KB</b></div>
      <div>🧠 Memory (peak)<br><b><?php echo $SH['memMB']; ?> MB</b></div>
      <div>💾 Last save<br><b><?php echo $SH['lastSave']?date('d M H:i:s',$SH['lastSave']):'—'; ?></b></div>
      <div>📦 Last backup<br><b><?php echo $SH['lastBackup']?date('d M H:i:s',$SH['lastBackup']):'—'; ?></b></div>
      <div>♻️ Last recovery<br><b><?php echo $SH['lastRecovery']?date('d M H:i:s',$SH['lastRecovery']):'kabhi zaroorat nahi padi ✅'; ?></b></div>
      <div>📋 Total records<br><b><?php echo number_format($SH['records']); ?></b></div>
      <div>🕒 Load / Save<br><b><?php echo $SH['loadMs']; ?> / <?php echo $SH['saveMs']; ?> ms</b></div>
    </div>
    <div style="font-size:10px;color:var(--mut);margin-top:8px">flock locking · atomic verify-writes · rotating .bak..bak4 · auto-recovery · caps (scans 1000 / feedback 500 / crash 500) · errors → <code>logs/error.log</code></div>
  </div>
  <?php } else { ?>
  <div class="card" id="cardStorage"><h3><span class="em">🗄️</span> Storage module</h3>
    <div style="font-size:12px;color:var(--mut)">⚠️ <b>json_storage.php</b> abhi upload nahi hui — panel legacy storage par chal raha hai (sab kaam karta hai). Upar "Upload &amp; update" me <b>json_storage.php (storage)</b> chunkar module upload karein — flock locking, rotating backups, auto-recovery aur error-log turant chalu ho jayenge.</div>
  </div>
  <?php } ?>

  </div><!-- /system -->

  <!-- ===== NEW PORTAL MODULES (client-side, D data se — backend untouched) ===== -->
  <div class="page" data-p="live">
    <!-- ===== MISSION CONTROL: header strip ===== -->
    <div class="card" style="display:flex;gap:14px;align-items:center;flex-wrap:wrap;background:linear-gradient(120deg,rgba(37,99,235,.12),rgba(124,58,237,.07));border-color:rgba(59,130,246,.35)">
      <div style="font-size:15px;font-weight:800">🛰 Mission Control</div>
      <span class="live"><span class="dot"></span> LIVE</span>
      <span style="color:var(--mut);font-size:10.5px">Auto-refresh: <b id="lvCount">30</b>s</span>
      <span style="color:var(--mut);font-size:10.5px">Server: <b id="lvResp">—</b></span>
      <span style="color:var(--mut);font-size:10.5px">🟢 <b id="lvOnN">0</b> online · 🖨 <b id="lvDevN">0</b> scanners</span>
      <div style="margin-left:auto;display:flex;gap:6px;flex-wrap:wrap" class="no-print">
        <button class="btn gray" onclick="location.reload()">🔄 Refresh</button>
        <button class="btn gray" onclick="jumpTo('overview','cardBroadcast')">📣 Broadcast</button>
        <a class="btn gray" style="text-decoration:none" href="?admin=1&dl=data">🗄 Backup</a>
        <button class="btn gray" onclick="lvCSV('all')">⬇ Export</button>
        <a class="btn gray" style="text-decoration:none" href="?admin=1&report=1" target="_blank">📑 Report</a>
      </div>
    </div>
    <!-- ===== KPI grid ===== -->
    <div class="kpis" id="liveKpis" style="margin-bottom:13px"></div>
    <!-- ===== version coverage + alerts + insights ===== -->
    <div class="grid3">
      <div class="card"><h3><span class="em">🔢</span> Version coverage</h3><div id="lvVer" class="skel"></div></div>
      <div class="card"><h3><span class="em">🚨</span> Alert Center <span class="no-print" style="margin-left:auto;display:flex;gap:4px" id="lvAlF">
        <button class="ufc on" data-s="all">All</button><button class="ufc" data-s="crit">Critical</button><button class="ufc" data-s="warn">Warning</button><button class="ufc" data-s="info">Info</button></span></h3>
        <div id="lvAlerts" class="skel" style="max-height:220px;overflow:auto"></div></div>
      <div class="card"><h3><span class="em">🤖</span> AI Insights</h3><div id="lvAI" class="skel" style="max-height:220px;overflow:auto"></div></div>
    </div>
    <!-- ===== live grid ===== -->
    <div class="card">
      <h3><span class="em">🖥</span> Live devices &amp; users <span style="color:var(--mut);font-weight:500;font-size:10px">(naam par click = poora profile)</span></h3>
      <div class="no-print" style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px;align-items:center">
        <input id="lvQ" placeholder="🔍 user / scanner / desh / IP / version…" style="flex:1;min-width:170px">
        <button class="ufc on" data-lf="all">Sab</button>
        <button class="ufc" data-lf="online">🟢 Online</button>
        <button class="ufc" data-lf="offline">🔴 Offline</button>
        <button class="ufc" data-lf="today">🖨 Aaj scan kiya</button>
        <button class="ufc" data-lf="oldver">⚠ Purana ver</button>
        <select id="lvScn" style="font-size:10.5px"><option value="">🖨 Sab scanner</option></select>
        <button class="btn gray" onclick="lvCSV('online')">⬇ Online CSV</button>
      </div>
      <div id="lvGrid" class="skel"></div>
    </div>
    <!-- ===== charts + heatmap + timeline ===== -->
    <div class="grid">
      <div class="card"><h3><span class="em">🕒</span> Scans — pichhle 24 ghante</h3><canvas id="lv24" height="130"></canvas></div>
      <div class="card"><h3><span class="em">📅</span> Scans — pichhle 14 din</h3><canvas id="lv14" height="130"></canvas></div>
    </div>
    <div class="grid">
      <div class="card"><h3><span class="em">🔥</span> Hafta × Ghanta heatmap <span style="color:var(--mut);font-weight:500;font-size:10px">— gehra = zyada scan</span></h3><div id="lvHeat" class="skel"></div></div>
      <div class="card"><h3><span class="em">🌍</span> Country distribution</h3><div id="lvGeo" class="skel"></div></div>
    </div>
    <div class="card"><h3><span class="em">📰</span> Live Activity Timeline</h3><div id="lvTime" class="skel" style="max-height:300px;overflow:auto"></div></div>
  </div>
  <div class="page" data-p="hw">
    <div class="sec"><span class="em">🖨</span> Devices — scanner ke hisaab se</div>
    <div class="kpis" id="hwKpis" style="margin-bottom:13px"></div>
    <div id="hwCards" class="grid skel"></div>
  </div>
  <div class="page" data-p="scans">
    <div class="sec"><span class="em">📄</span> Scans — data grid (last 300)</div>
    <div class="kpis" id="sgKpis" style="margin-bottom:13px"></div>
    <div class="card">
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:9px;align-items:center">
        <input id="sgQ" placeholder="🔍 user / country / profile se dhoondo…" style="flex:1;min-width:160px">
        <select id="sgRange"><option value="0">Sab</option><option value="1">Aaj</option><option value="7">7 din</option><option value="30">30 din</option></select>
        <select id="sgDpi"><option value="">DPI: sab</option></select>
        <select id="sgCol"><option value="">Rang: sab</option></select>
        <select id="sgSrc"><option value="">Srot: sab</option><option value="feeder">Feeder</option><option value="glass">Glass</option></select>
        <button class="btn gray" onclick="sgCSV()">⬇ CSV</button>
      </div>
      <div style="color:var(--mut);font-size:10.5px;margin-bottom:7px">💡 Kisi bhi row par CLICK karo — us scan ka poora byora khul jayega (scanner, profile, samay…)</div>
      <div id="sgGrid" class="skel"></div>
      <div id="sgPag" style="margin-top:9px;display:flex;gap:6px;align-items:center;justify-content:center"></div>
    </div>
  </div>
  <div class="page" data-p="reports">
    <div class="sec"><span class="em">📑</span> Reports &amp; Export</div>
    <div class="grid">
      <div class="card"><h3><span class="em">📄</span> Branded report (PDF)</h3>
        <div style="color:var(--mut);margin-bottom:8px">Ek-page ka summary — print dialog se "Save as PDF".</div>
        <a class="btn" href="?admin=1&report=1" target="_blank">📄 Report kholo / PDF banao</a></div>
      <div class="card"><h3><span class="em">📊</span> CSV / JSON export</h3>
        <div style="color:var(--mut);margin-bottom:8px">Excel me kholne layak CSV, ya poora JSON backup.</div>
        <a class="rbtn d" href="?admin=1&export=days">Daily CSV</a>
        <a class="rbtn d" href="?admin=1&export=months">Monthly CSV</a>
        <a class="rbtn d" href="?admin=1&export=users">Users CSV</a>
        <a class="rbtn d" href="?admin=1&export=countries">Country CSV</a>
        <a class="rbtn d" href="?admin=1&export=scans">Scans CSV</a>
        <a class="rbtn d" href="?admin=1&export=crashes">Crash CSV</a>
        <a class="rbtn d" href="?admin=1&export=feedback">Feedback CSV</a>
        <a class="rbtn d" href="?admin=1&export=json">Backup (JSON)</a>
        <a class="rbtn d" href="?admin=1&dl=data">stats.json (dated)</a></div>
      <div class="card"><h3><span class="em">🖨</span> Print dashboard</h3>
        <div style="color:var(--mut);margin-bottom:8px">Poora dashboard printer/PDF par.</div>
        <button class="btn" onclick="window.print()">🖨 Print</button></div>
      <div class="card"><h3><span class="em">📄</span> Scan grid CSV</h3>
        <div style="color:var(--mut);margin-bottom:8px">"Scans" module ki filtered list CSV me.</div>
        <button class="btn" onclick="jumpToPage('scans');setTimeout(sgCSV,300)">⬇ Scans CSV</button></div>
    </div>
  </div>
  <div class="page" data-p="health">
    <div class="sec"><span class="em">❤️</span> System Health</div>
    <div class="kpis" id="healthKpis" style="margin-bottom:13px"></div>
    <div class="card"><h3><span class="em">🗄️</span> JSON storage layer</h3><div id="healthStore" class="skel"></div></div>
    <div class="grid" style="margin-top:13px">
      <div class="card"><h3><span class="em">📦</span> Backups (server par asli files)</h3><div id="healthBk" class="skel"></div></div>
      <div class="card"><h3><span class="em">📋</span> error.log (aakhri lines)</h3><div id="healthLog" class="skel"></div></div>
    </div>
  </div>
  <div class="page" data-p="logs">
    <div class="sec"><span class="em">📋</span> Activity Logs</div>
    <div class="grid">
      <div class="card"><h3><span class="em">🔐</span> Admin logins</h3><div id="logLogins" class="skel"></div></div>
      <div class="card"><h3><span class="em">🛠</span> Admin actions (audit)</h3><div id="logAudit" class="skel"></div></div>
      <div class="card"><h3><span class="em">🚨</span> Failed login attempts (security)</h3><div id="logFails" class="skel"></div></div>
    </div>
  </div>

  
  <div class="foot">Server: PHP <?php echo htmlspecialchars($S['srv']); ?> · <?php echo htmlspecialchars($S['time']); ?> · ApneSoftware.com</div>
</div>
<script>
var D=<?php echo $J; ?>;
document.getElementById('tm').textContent=D.time;
document.getElementById('rt').textContent='server '+D.respMs+'ms';
// Windows emoji-flags nahi dikhata ("us US" jaisa kachra dikhta tha) —
// isliye saaf text-badge (IN / US / AE) har jagah kaam karta hai.
function flag(cc){ if(!cc||cc.length!==2) return '';
  return '<span style="display:inline-block;padding:0 5px;border-radius:5px;background:var(--line);color:var(--fg);font-size:9.5px;font-weight:800;letter-spacing:.4px;vertical-align:middle">'+String(cc).toUpperCase()+'</span>'; }
function fmt(n){ return (n||0).toLocaleString(); }
function esc(s){ return (s==null?'':(''+s)).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];}); }
function ago(ts){ if(!ts)return '—'; var s=Math.floor(Date.now()/1000)-ts; if(s<60)return s+'s'; if(s<3600)return Math.floor(s/60)+'m'; if(s<86400)return Math.floor(s/3600)+'h'; return Math.floor(s/86400)+'d'; }

// theme
function toggleTh(){ var lt=document.documentElement.getAttribute('data-th')==='light'; try{localStorage.setItem('anth',lt?'dark':'light');}catch(e){} location.reload(); }
(function(){ var th='dark'; try{ if(localStorage.getItem('anth')==='light') th='light'; }catch(e){} if(th==='light'){ document.documentElement.setAttribute('data-th','light'); var b=document.getElementById('thbtn'); if(b)b.textContent='🌙'; } })();

// palette (validated categorical) + Chart.js theme
var DKT=document.documentElement.getAttribute('data-th')!=='light';   // dark-first
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
var FEATLBL={ocr:'OCR (text)',compress:'Compress',merge:'Merge',split:'Split page',sign:'Signature',stamp:'Stamp',password:'Password',watermark:'Watermark',whatsapp:'WhatsApp share',email:'Email share',print:'Print',import:'Import',phoneimport:'Phone photo',idcard:'ID-card crop',rename:'Rename',donate:'Donate click',refer:'Share app',phonescan:'Phone scan',save:'PDF save',search:'File search'};

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
var SENT={pos:{e:'😊',c:PAL.green,t:'Khush'},neg:{e:'😟',c:PAL.red,t:'Naraz'},req:{e:'💡',c:PAL.orange,t:'Farmaish'},neu:{e:'😐',c:'var(--mut)',t:'Neutral'}};
(function(){var s=D.sentiment||{};var parts=['pos','neg','req','neu'].filter(function(k){return s[k];}).map(function(k){return '<span style="color:'+SENT[k].c+'">'+SENT[k].e+' '+s[k]+'</span>';});document.getElementById('sentsum').innerHTML=parts.length?('— '+parts.join(' · ')):'';})();
var CANNED=[]; try{var _rc=JSON.parse(D.rconfigStr||'{}'); CANNED=_rc.canned||[];}catch(e){}
function cannedSel(fi){ if(!CANNED.length) return ''; return '<select class="no-print" style="font-size:10px;max-width:120px" onchange="var f=this.closest(\'form\');if(f&&this.value)f.reply.value=this.value"><option value="">📋 Ready jawab…</option>'+CANNED.map(function(c){return '<option>'+esc(c)+'</option>';}).join('')+'</select>'; }
document.getElementById('fb').innerHTML=(D.feedback&&D.feedback.length)?D.feedback.map(function(f){var se=SENT[f.sent||'neu'];var resolved=(f.status==='resolved');return '<div style="border-bottom:1px solid var(--line);padding:6px 0'+(resolved?';opacity:.55':'')+'"><b>'+('★'.repeat(f.rating||0)||'—')+'</b> '+esc(f.name||'—')+' <span style="color:var(--mut);font-size:11px">'+(f.v?'v'+f.v:'')+'</span> <span title="'+se.t+'" style="font-size:11px;color:'+se.c+'">'+se.e+'</span> <form method="post" class="no-print" style="display:inline"><input type="hidden" name="act" value="fstatus"><input type="hidden" name="fi" value="'+f.fi+'"><button class="btn" style="padding:1px 7px;font-size:10px;background:'+(resolved?'var(--accent2)':'var(--line)')+';color:'+(resolved?'#fff':'var(--fg)')+'">'+(resolved?'✅ Resolved':'⏳ Open')+'</button></form><div>'+esc(f.msg||'')+'</div>'+(f.reply?'<div style="margin:4px 0 2px;padding:5px 8px;background:var(--card2);border-left:3px solid var(--accent);border-radius:0 6px 6px 0;font-size:12px">↩️ <b>Aapne kaha:</b> '+esc(f.reply)+'</div>':'<form method="post" class="no-print" style="margin-top:4px;display:flex;gap:4px;align-items:center;flex-wrap:wrap"><input type="hidden" name="act" value="freply"><input type="hidden" name="fi" value="'+f.fi+'"><input name="reply" placeholder="Jawab (user ki app me dikhega)…" style="flex:1;min-width:120px;font-size:11px">'+cannedSel(f.fi)+'<button class="btn" style="padding:4px 8px">↩️</button></form>')+'</div>';}).join(''):'<div style="color:var(--mut);font-size:12px">— abhi koi feedback nahi —</div>';
document.getElementById('cr').innerHTML=(D.crashes&&D.crashes.length)?'<table>'+D.crashes.map(function(c){return '<tr><td>💥 '+esc((c.err||'').slice(0,60))+'</td><td style="text-align:right;color:var(--mut);white-space:nowrap">v'+(c.v||'?')+' · '+ago(c.t)+'</td></tr>';}).join('')+'</table>':'<div style="color:var(--mut);font-size:12px">— koi crash nahi 🎉 —</div>';

// churn-risk / at-risk (13)
document.getElementById('atr').innerHTML=(D.atRisk&&D.atRisk.length)?'<table>'+D.atRisk.map(function(u){return '<tr><td>😟 '+esc(u.name)+'</td><td style="text-align:right"><b>'+u.scans+'</b> scans · '+u.days+' din se gayab</td></tr>';}).join('')+'</table>':'<div style="color:var(--mut);font-size:12px">— koi risk par nahi 🎉 —</div>';

// version quality / crash-rate (15)
document.getElementById('vq').innerHTML=(D.versionQuality&&D.versionQuality.length)?'<table><tr><th>Version</th><th style="text-align:right">Users</th><th style="text-align:right">Crashes</th><th style="text-align:right">Rate</th></tr>'+D.versionQuality.map(function(v){var bad=v.crashes>0&&v.rate>=0.3;return '<tr><td>v'+esc(v.v)+'</td><td style="text-align:right">'+fmt(v.users)+'</td><td style="text-align:right">'+v.crashes+'</td><td style="text-align:right;font-weight:700;color:'+(bad?PAL.red:(v.crashes?PAL.orange:PAL.green))+'">'+v.rate+'</td></tr>';}).join('')+'</table>':'<div style="color:var(--mut);font-size:12px">— data nahi —</div>';

// feature correlation (16)
document.getElementById('fcorr').innerHTML=(D.featCorr&&D.featCorr.length)?'<table>'+D.featCorr.map(function(c){var A=(FEATLBL[c.a]||c.a),B=(FEATLBL[c.b]||c.b);return '<tr><td><b>'+esc(A)+'</b> → '+esc(B)+'</td><td style="text-align:right"><b>'+c.pct+'%</b> <span style="color:var(--mut)">('+c.both+')</span></td></tr>';}).join('')+'</table>':'<div style="color:var(--mut);font-size:12px">— abhi kaafi data nahi —</div>';

// prefill remote-config editors from live rconfig
(function(){ var rc={}; try{rc=JSON.parse(D.rconfigStr||'{}');}catch(e){}
  try{ if(rc.faq) document.getElementById('faqbox').value=JSON.stringify(rc.faq,null,1); }catch(e){}
  try{ document.getElementById('donbox').value=rc.donate_url||''; }catch(e){}
  try{ if(rc.tips&&rc.tips.length) document.getElementById('tipbox').value=rc.tips.join('\n'); }catch(e){}
  try{ if(rc.flags) document.getElementById('flagbox').value=JSON.stringify(rc.flags); }catch(e){}
  try{ if(rc.canned&&rc.canned.length) document.getElementById('cannedbox').value=rc.canned.join('\n'); }catch(e){}
  try{ document.getElementById('chlogbox').value=rc.changelog||''; }catch(e){}
  try{ if(rc.news&&rc.news.length) document.getElementById('newsbox').value=rc.news.join('\n'); }catch(e){}
})();

// (27) crash groups
window.crashFix=function(h,un){ var f=document.createElement('form'); f.method='post'; f.style.display='none';
  f.innerHTML='<input name="act" value="'+(un?'unfixcrash':'fixcrash')+'"><input name="sig" value="'+h+'">';
  document.body.appendChild(f); f.submit(); };
document.getElementById('cg').innerHTML=(D.crashGroups&&D.crashGroups.length)?'<table>'+D.crashGroups.map(function(g){
  return '<tr'+(g.fixed?' style="opacity:.55"':'')+'><td>'+(g.fixed?'✅':'💥')+' '+esc(g.sample||g.sig)
    +(g.vers?' <span style="color:var(--mut);font-size:10px">v'+esc(g.vers)+'</span>':'')
    +(g.fixed?' <span style="color:var(--ok);font-size:10px;font-weight:700">RESOLVED</span>':'')
    +'</td><td style="text-align:right;white-space:nowrap"><b>×'+g.count+'</b> · '+ago(g.last)
    +(g.h?' <button class="btn gray" style="padding:2px 7px;font-size:10px" onclick="crashFix(\''+g.h+'\','+(g.fixed?1:0)+')">'+(g.fixed?'Reopen':'✔ Resolve')+'</button>':'')
    +'</td></tr>';}).join('')+'</table>':'<div style="color:var(--mut);font-size:12px">— koi crash nahi 🎉 —</div>';

// (30) referral leaderboard
document.getElementById('rf').innerHTML=(D.refLeaders&&D.refLeaders.length)?'<table>'+D.refLeaders.map(function(u,i){var m=(i<3?['🥇','🥈','🥉'][i]:(i+1)+'.');return '<tr><td>'+m+' '+esc(u.name)+(u.cc?' '+flag(u.cc):'')+'</td><td style="text-align:right"><b>'+u.n+'</b> share</td></tr>';}).join('')+'</table>':'<div style="color:var(--mut);font-size:12px">— abhi kisi ne share nahi kiya —</div>';

// (32) top users leaderboard (admin view; public opt-in v141 me)
(function(){ var ul=(D.userList||[]).slice().filter(function(u){return u.scans>0;}).sort(function(a,b){return b.scans-a.scans;}).slice(0,12);
  document.getElementById('lb').innerHTML=ul.length?'<table>'+ul.map(function(u,i){var m=(i<3?['🥇','🥈','🥉'][i]:(i+1)+'.');return '<tr><td>'+m+' '+esc(u.name)+(u.country?' '+flag(u.country):'')+'</td><td style="text-align:right"><b>'+fmt(u.scans)+'</b></td></tr>';}).join('')+'</table>':'<div style="color:var(--mut);font-size:12px">— data nahi —</div>'; })();

// (26) server health — request load chart + data-size trend
(function(){
  var rh=D.reqHours||{}, ks=Object.keys(rh); var lab=ks.map(function(k){return k.slice(11)+'h';}), val=ks.map(function(k){return rh[k];});
  var sd=D.sizeDaily||{}, sk=Object.keys(sd); var last=sk.length?sd[sk[sk.length-1]]:(D.fileKB||0);
  var hpk=Math.max.apply(null,val.concat([0]));
  document.getElementById('hl').innerHTML='<table><tr><td>📦 Data file</td><td style="text-align:right"><b>'+last+' KB</b></td></tr>'+
    '<tr><td>⚡ Peak req/ghanta</td><td style="text-align:right"><b>'+hpk+'</b></td></tr>'+
    '<tr><td>🕒 Response abhi</td><td style="text-align:right"><b>'+D.respMs+' ms</b></td></tr></table>';
  var c=document.getElementById('hlc'); if(c&&window.Chart) mkChart(c,{type:'bar',data:{labels:lab,datasets:[{data:val,backgroundColor:PAL.aqua}]},options:CO});
})();

// (3) audit log
document.getElementById('au').innerHTML=(D.auditLog&&D.auditLog.length)?'<table>'+D.auditLog.map(function(a){return '<tr><td><b>'+esc(a.act)+'</b>'+(a.det?' <span style="color:var(--mut)">'+esc(a.det)+'</span>':'')+'</td><td style="text-align:right;color:var(--mut);white-space:nowrap">'+esc(a.ip)+' · '+ago(a.t)+'</td></tr>';}).join('')+'</table>':'<div style="color:var(--mut);font-size:12px">— abhi koi action nahi —</div>';

// (15) retention cohort heatmap grid
(function(){ var R=D.retention||[]; if(!R.length){ document.getElementById('retgrid').innerHTML='<div style="color:var(--mut);font-size:12px">— abhi kaafi data nahi —</div>'; return; }
  function hcol(p){ if(p<=0)return 'var(--line)'; var a=0.15+0.85*(p/100); return 'rgba(27,175,122,'+a.toFixed(2)+')'; }
  var h='<table style="border-collapse:separate;border-spacing:2px;font-size:10px"><tr><th style="text-align:left">Cohort (naye)</th>';
  for(var w=0;w<=7;w++) h+='<th style="width:34px;text-align:center;color:var(--mut)">W'+w+'</th>';
  h+='</tr>';
  R.forEach(function(r){ h+='<tr><td style="white-space:nowrap"><b>'+esc(r.label)+'</b> <span style="color:var(--mut)">('+r.size+')</span></td>';
    r.pct.forEach(function(p){ h+='<td style="text-align:center;padding:4px 2px;border-radius:4px;color:'+(p>45?'#fff':'var(--fg2)')+';background:'+hcol(p)+'">'+(p>0?p:'·')+'</td>'; });
    h+='</tr>'; });
  document.getElementById('retgrid').innerHTML=h+'</table>';
})();

// (18) feature adoption over time (multi-line)
(function(){ var FA=D.featAdopt||{days:[],series:{}}; var el=document.getElementById('featAdopt'); if(!el||!window.Chart) return;
  var days=FA.days||[]; var series=FA.series||{}; var keys=Object.keys(series);
  if(!keys.length){ el.parentNode.innerHTML+='<div style="color:var(--mut);font-size:12px">— abhi kaafi data nahi (feature use hote hi bharега) —</div>'; return; }
  var cols=[PAL.blue,PAL.orange,PAL.aqua,PAL.violet,PAL.magenta];
  var ds=keys.map(function(k,i){return {label:(FEATLBL[k]||k),data:series[k],borderColor:cols[i%cols.length],backgroundColor:cols[i%cols.length],borderWidth:2,tension:.35,pointRadius:0,fill:false};});
  mkChart(el,{type:'line',data:{labels:days.map(function(x){return x.slice(5);}),datasets:ds},options:{plugins:{legend:{display:true,position:'bottom',labels:{boxWidth:10,font:{size:10}}}},scales:CO.scales}});
})();

// ---- Rename analytics ----
(function(){ var R=D.rename||{};
  document.getElementById('rnKpis').innerHTML=
    kpi('✏️',fmt(R.total),'Total renames')+
    kpi('👥',fmt(R.users),'Rename karne waale','p')+
    kpi('📥',(R.adopt||0)+'%','Adoption (users)','o')+
    kpi('🔁',(R.perUser||0),'Per user','y')+
    kpi('📅',fmt(R.today),'Aaj')+
    kpi('🗓️',fmt(R.week),'Is hafte')+
    kpi('📆',fmt(R.month),'Is mahine')+
    kpi('💡',(R.metaN>0?(R.pickRate||0)+'%':'—'),'Suggestion pick-rate','g');
  // trend chart
  var tr=R.trend||[]; var el=document.getElementById('rnTrend');
  if(el&&window.Chart) mkChart(el,{type:'line',data:{labels:tr.map(function(x){return x[0].slice(5);}),datasets:[{data:tr.map(function(x){return x[1];}),borderColor:PAL.violet,backgroundColor:function(c){return grad(c,PAL.violet);},borderWidth:2,fill:true,tension:.35,pointRadius:0}]},options:CO});
  // leaders
  var L=R.leaders||[]; document.getElementById('rnLeaders').innerHTML=L.length?'<table>'+L.map(function(u,i){var m=(i<3?['🥇','🥈','🥉'][i]:(i+1)+'.');return '<tr><td>'+m+' '+esc(u.name)+(u.cc?' '+flag(u.cc):'')+'</td><td style="text-align:right"><b>'+fmt(u.n)+'</b></td></tr>';}).join('')+'</table>':'<div style="color:var(--mut);font-size:12px">— abhi kisi ne rename nahi kiya —</div>';
  // by country
  var bc=R.byCountry||{}; var bck=Object.keys(bc).sort(function(a,b){return bc[b]-bc[a];}).slice(0,10);
  document.getElementById('rnByCo').innerHTML=bck.length?'<table>'+bck.map(function(cc){return '<tr><td>'+flag(cc)+' '+cc+'</td><td style="text-align:right"><b>'+fmt(bc[cc])+'</b></td></tr>';}).join('')+'</table>':'<div style="color:var(--mut);font-size:12px">— data nahi —</div>';
  // habits (needs v141 meta)
  if(R.metaN>0){ document.getElementById('rnHabits').innerHTML='<table>'+
    '<tr><td>💡 Suggestion se chuna</td><td style="text-align:right"><b>'+(R.pickRate||0)+'%</b></td></tr>'+
    '<tr><td>🔤 Average lambaai</td><td style="text-align:right"><b>'+(R.avgLen||0)+' akshar</b></td></tr>'+
    '<tr><td>📝 Average shabd</td><td style="text-align:right"><b>'+(R.avgWords||0)+'</b></td></tr>'+
    '<tr><td>🔢 Number waale naam</td><td style="text-align:right"><b>'+(R.pctNum||0)+'%</b></td></tr>'+
    '<tr><td>📅 Date waale naam</td><td style="text-align:right"><b>'+(R.pctDate||0)+'%</b></td></tr></table>'+
    '<div style="color:var(--mut);font-size:10px;margin-top:6px">'+fmt(R.metaN)+' renames ka data · koi asli naam nahi (sirf ginti)</div>';
  } else { document.getElementById('rnHabits').innerHTML='<div style="color:var(--mut);font-size:12px">Ye deep data (pick-rate, lambaai, patterns) <b>v141</b> update ke baad bharna shuru hoga — app privacy-safe meta bhejegी (koi asli naam nahi).</div>'; }
})();

// (20) custom report builder
var RBM=[['total','Total scans'],['today','Aaj'],['weekTotal','Is hafte'],['monthTotal','Is mahine'],['users','Total users'],['newToday','Aaj naye'],['dau','DAU'],['wau','WAU'],['mau','MAU'],['powerUsers','Power users'],['avgRating','Avg rating'],['online','Online abhi'],['imports','Imports'],['prints','Prints'],['dailyAvg','Daily avg'],['forecastMonth','Forecast (mahina)']];
(function(){ var p=document.getElementById('rbpick'); if(!p)return; p.innerHTML=RBM.map(function(m,i){return '<label style="font-size:11px;display:flex;align-items:center;gap:4px"><input type="checkbox" class="rbchk" value="'+m[0]+'" '+(i<8?'checked':'')+'>'+m[1]+'</label>';}).join(''); })();
function buildReport(){ var sel=[].filter.call(document.querySelectorAll('.rbchk'),function(c){return c.checked;}).map(function(c){return c.value;});
  var lbl={}; RBM.forEach(function(m){lbl[m[0]]=m[1];});
  var h='<div style="font-weight:700;margin-bottom:6px">📊 ApneScan report — '+D.time+'</div><div class="kpis">';
  sel.forEach(function(k){ h+='<div class="kpi"><div class="tx"><div class="n">'+fmt(D[k]||0)+'</div><div class="l">'+esc(lbl[k]||k)+'</div></div></div>'; });
  h+='</div><button class="btn no-print" style="margin-top:8px" onclick="window.print()">🖨 Print / PDF</button>';
  document.getElementById('rbout').innerHTML=h;
}

// (10) saved segments (per-browser, localStorage)
function loadSegs(){ try{ return JSON.parse(localStorage.getItem('an_segs')||'[]'); }catch(e){ return []; } }
function renderSegs(){ var segs=loadSegs(); var bar=document.getElementById('segbar'); if(!bar)return;
  bar.innerHTML=segs.length?segs.map(function(s,i){return '<span class="tag" style="cursor:pointer" onclick="applySeg('+i+')">🔍 '+esc(s.q||'(sab)')+'</span><span onclick="delSeg('+i+')" style="cursor:pointer;color:var(--mut);margin-left:-3px;margin-right:4px" title="hatao">✕</span>';}).join(''):'<span style="font-size:11px;color:var(--mut)">— koi nahi —</span>';
}
function saveSeg(){ var q=(document.getElementById('usearch').value||'').trim(); var segs=loadSegs(); segs.push({q:q}); segs=segs.slice(-12); try{localStorage.setItem('an_segs',JSON.stringify(segs));}catch(e){} renderSegs(); }
function applySeg(i){ var s=loadSegs()[i]; if(!s)return; document.getElementById('usearch').value=s.q||''; renderUsers(); }
function delSeg(i){ var segs=loadSegs(); segs.splice(i,1); try{localStorage.setItem('an_segs',JSON.stringify(segs));}catch(e){} renderSegs(); }
renderSegs();

// records + admin logins
document.getElementById('rc').innerHTML='<table>'+
 '<tr><td>🏔 Peak online (all-time)</td><td style="text-align:right"><b>'+D.peakAll+'</b></td></tr>'+
 '<tr><td>📅 Best single day</td><td style="text-align:right"><b>'+D.bestDay+'</b></td></tr>'+
 '<tr><td>🕐 This hour</td><td style="text-align:right"><b>'+D.hour+'</b></td></tr>'+
 '<tr><td>🆕 Latest version</td><td style="text-align:right"><b>v'+D.latestVersion+'</b></td></tr></table>';
document.getElementById('al').innerHTML='<table>'+(D.adminLogins||[]).map(function(t){return '<tr><td>🔒 '+t[0]+'</td><td style="text-align:right;color:var(--mut)">'+t[1]+'</td></tr>';}).join('')+'</table>';

// ---- user table ----
var sortKey='scans', sortDir=-1;
// ================= ENTERPRISE USER MANAGEMENT =================
// Health score (asli data se): crashes, purana version, offline duration, activity
function uHealth(u){
  var sc=100, why=[];
  var crashes=(D.crashes||[]).filter(function(c){return c.client&&u.id&&String(c.client)===String(u.id).slice(0,40);}).length;
  if(crashes){ sc-=Math.min(30,crashes*10); why.push(crashes+' crash'); }
  if(u.version&&D.latestVersion&&String(u.version).trim()!==String(D.latestVersion).trim()){ sc-=15; why.push('purana v'+u.version); }
  var offd=u.last?Math.floor((Date.now()/1000-u.last)/86400):999;
  if(offd>=30){sc-=25;why.push(offd+'d offline');} else if(offd>=14){sc-=15;why.push(offd+'d offline');} else if(offd>=7){sc-=7;}
  var dm=u.daysMap||{},w=0,dt=new Date();
  for(var i=0;i<7;i++){var k=new Date(dt.getTime()-i*86400000).toISOString().slice(0,10);w+=parseInt(dm[k])||0;}
  if(w===0&&offd<7)sc-=5;
  if(u.blocked){sc=Math.min(sc,20);why.push('blocked');}
  return {s:Math.max(5,sc), why:why.join(' · ')||'sab theek', week:w};
}
function uToday(u){ var k=new Date().toISOString().slice(0,10); return parseInt((u.daysMap||{})[k])||0; }
function uStatus(u){
  if(u.blocked) return '<span class="stb b">🚫 Blocked</span>';
  var out=u.online?'<span class="stb g">🟢 Online</span>'
    :((Date.now()/1000-(u.last||0))<1800?'<span class="stb y">🟡 Idle</span>':'<span class="stb r">🔴 Offline</span>');
  if(u.version&&D.latestVersion&&String(u.version).trim()!==String(D.latestVersion).trim()) out+=' <span class="stb w">⚠ Update</span>';
  if((u.tags||'').toLowerCase().indexOf('vip')>-1||(u.tags||'').toLowerCase().indexOf('premium')>-1) out+=' <span class="stb p">⭐ VIP</span>';
  return out;
}
function uAvatar(u,big){
  var nm=(u.name&&u.name!=='—')?u.name:'?';
  var ini=nm.trim().split(/\s+/).slice(0,2).map(function(w){return (w[0]||'').toUpperCase();}).join('')||'?';
  var cols=['#2563EB','#7C3AED','#0D9488','#DB2777','#EA580C','#16A34A','#0891B2','#9333EA'];
  var hsh=0; for(var i=0;i<nm.length;i++)hsh=(hsh*31+nm.charCodeAt(i))>>>0;
  return '<span class="uav'+(big?' big':'')+'" style="background:'+cols[hsh%cols.length]+'">'+esc(ini)+(u.online?'<span class="on"></span>':'')+'</span>';
}
function uCSV(rows){
  var csv='name,status,scanner,total_scans,today,last_seen,joined,version,country,region,method,health\n'
    +rows.map(function(u){var h=uHealth(u);
      return '"'+String(u.name).replace(/"/g,'""')+'",'+(u.online?'online':'offline')+',"'+String(u.model||'').replace(/"/g,'""')+'",'
      +u.scans+','+uToday(u)+','+(u.last?new Date(u.last*1000).toISOString():'')+','+(u.first?new Date(u.first*1000).toISOString().slice(0,10):'')
      +','+(u.version||'')+','+(u.country||'')+',"'+String(u.region||'').replace(/"/g,'""')+'",'+(u.method||'')+','+h.s;}).join('\n');
  var a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));a.download='apnescan-users.csv';a.click();
  if(window.showToast)showToast('⬇ Users CSV taiyaar','ok');
}
var _uf='all';
(function(){ var bar=document.getElementById('ufbar'); if(!bar)return;
  [].forEach.call(bar.querySelectorAll('.ufc'),function(b){ b.onclick=function(){
    [].forEach.call(bar.querySelectorAll('.ufc'),function(x){x.classList.remove('on');});
    b.classList.add('on'); _uf=b.getAttribute('data-f'); renderUsers(); }; });
  var cs=document.getElementById('ufCountry'); var seen={};
  (D.userList||[]).forEach(function(u){ if(u.country&&!seen[u.country]){seen[u.country]=1;
    var o=document.createElement('option'); o.value=u.country; o.textContent='🌍 '+u.country; cs.appendChild(o);} });
  cs.onchange=renderUsers;
})();
function renderUsers(){
  var q=(document.getElementById('usearch').value||'').toLowerCase();
  var cn=(document.getElementById('ufCountry')||{}).value||'';
  var med=0; (D.userList||[]).forEach(function(u){med=Math.max(med,u.scans||0);});
  var rows=(D.userList||[]).filter(function(u){
    if(q&&(u.name+' '+u.country+' '+(u.region||'')+' '+u.version+' '+u.method+' '+u.tags+' '+u.model).toLowerCase().indexOf(q)<0)return false;
    if(cn&&u.country!==cn)return false;
    var offd=u.last?Math.floor((Date.now()/1000-u.last)/86400):999;
    if(_uf==='online'&&!u.online)return false;
    if(_uf==='offline'&&u.online)return false;
    if(_uf==='today'&&uToday(u)<=0)return false;
    if(_uf==='oldver'&&!(u.version&&D.latestVersion&&u.version!==D.latestVersion))return false;
    if(_uf==='inactive'&&offd<14)return false;
    if(_uf==='heavy'&&(u.scans||0)<Math.max(10,med*0.4))return false;
    if(_uf==='attention'&&uHealth(u).s>=80)return false;
    return true;
  });
  rows.forEach(function(u){ u._h=uHealth(u).s; u._t=uToday(u); });
  rows.sort(function(a,b){ var x=a[sortKey],y=b[sortKey]; if(typeof x==='string'){return sortDir*String(x).localeCompare(String(y));} return sortDir*((x||0)-(y||0)); });
  document.getElementById('ucount').textContent=rows.length;
  var COLS=[['name','User'],['scans','Scans'],['_t','Aaj'],['last','Last seen'],['first','Joined'],['version','Ver'],['country','Location'],['method','Method'],['_h','Health']];
  var h='<tr><th class="no-print" style="width:22px"><input type="checkbox" id="uall" title="Sab chuno"></th>'
    +COLS.map(function(c){return '<th data-k="'+c[0]+'">'+c[1]+(sortKey===c[0]?(sortDir<0?' ▼':' ▲'):'')+'</th>';}).join('')
    +'<th class="no-print">Actions</th></tr>';
  rows.forEach(function(u,i){
    var hh=u._h, hc=hh>=80?'var(--ok)':(hh>=60?'var(--warn)':'var(--bad)');
    h+='<tr data-i="'+i+'" style="'+(u.blocked?'opacity:.45':'')+'">'
      +'<td class="no-print" style="text-align:center"><input type="checkbox" class="ubox" data-id="'+esc(u.id)+'"></td>'
      +'<td style="cursor:pointer;min-width:150px" data-open="'+i+'"><div style="display:flex;gap:8px;align-items:center">'+uAvatar(u)
        +'<div><div style="font-weight:700">'+esc(u.name)+'</div><div style="margin-top:1px">'+uStatus(u)
        +(u.tags?' '+u.tags.split(',').slice(0,2).map(function(t){return '<span class="tag">'+esc(t.trim())+'</span>';}).join(''):'')+'</div></div></div></td>'
      +'<td><b>'+fmt(u.scans)+'</b><div style="color:var(--mut);font-size:9px">'+esc((u.model||'').slice(0,18)||'—')+'</div></td>'
      +'<td>'+(u._t?'<b style="color:var(--ok)">'+u._t+'</b>':'<span style="color:var(--mut)">0</span>')+'</td>'
      +'<td>'+ago(u.last)+'</td>'
      +'<td>'+(u.first?new Date(u.first*1000).toISOString().slice(0,10):'—')+'</td>'
      +'<td>'+(u.version?('v'+u.version+(D.latestVersion&&u.version!==D.latestVersion?' ⚠':'')):'—')+'</td>'
      +'<td>'+(u.country?flag(u.country)+' '+u.country:'—')+(u.region?'<div style="color:var(--mut);font-size:9px">'+esc(u.region)+'</div>':'')+'</td>'
      +'<td>'+({escl:'🌐 Network',wia:'🔌 USB',twain:'🔌 TWAIN',naps2:'NAPS2'}[u.method]||u.method||'—')+'</td>'
      +'<td title="'+esc(uHealth(u).why)+'"><span class="hbar"><i style="width:'+hh+'%;background:'+hc+'"></i></span> <b style="color:'+hc+';font-size:10px">'+hh+'%</b></td>'
      +'<td class="no-print"><span class="uact">'
        +'<button title="Profile" data-open="'+i+'">👁</button>'
        +'<button title="Message" data-msg="'+i+'">📨</button>'
        +'<button title="Export CSV" data-csv="'+i+'">⬇</button></span></td>'
      +'</tr>';
  });
  var t=document.getElementById('utable'); t.innerHTML=h;
  [].forEach.call(t.querySelectorAll('th[data-k]'),function(th){ th.style.cursor='pointer'; th.onclick=function(){ var k=th.getAttribute('data-k'); if(sortKey===k)sortDir*=-1; else {sortKey=k;sortDir=(k==='name'||k==='country'||k==='version'||k==='method')?1:-1;} renderUsers(); }; });
  [].forEach.call(t.querySelectorAll('[data-open]'),function(td){ td.onclick=function(){ showUser(rows[+td.getAttribute('data-open')]); }; });
  [].forEach.call(t.querySelectorAll('[data-msg]'),function(b){ b.onclick=function(e){ e.stopPropagation(); showUser(rows[+b.getAttribute('data-msg')]); setTimeout(function(){var f=document.querySelector('#umbody input[name=msg]'); if(f)f.focus();},150); }; });
  [].forEach.call(t.querySelectorAll('[data-csv]'),function(b){ b.onclick=function(e){ e.stopPropagation(); uCSV([rows[+b.getAttribute('data-csv')]]); }; });
  [].forEach.call(t.querySelectorAll('.ubox'),function(b){ b.onchange=bulkSync; });
  var ua=document.getElementById('uall'); if(ua) ua.onchange=function(){ [].forEach.call(t.querySelectorAll('.ubox'),function(b){ b.checked=ua.checked; }); bulkSync(); };
  bulkSync();
}
function bulkSync(){ var ids=[].filter.call(document.querySelectorAll('.ubox'),function(b){return b.checked;}).map(function(b){return b.getAttribute('data-id');}); document.getElementById('bulkids').value=ids.join(','); document.getElementById('bulkn').textContent=ids.length; document.getElementById('bulkform').style.display=ids.length?'flex':'none'; }
(function(){ var ba=document.getElementById('bulkba'); if(ba) ba.onchange=function(){ document.getElementById('bulktags').style.display=ba.value==='tag'?'':'none'; document.getElementById('bulkmsg').style.display=ba.value==='msg'?'':'none'; }; })();
function bulkGo(){ var n=document.getElementById('bulkids').value.split(',').filter(Boolean).length; if(!n){alert('Pehle users chuno');return false;} return confirm(n+' users par ye action lagaayein?'); }
// ---------------- PROFILE PAGE (user details) ----------------
function showUser(u){
  if(!u) return;
  var mth={escl:'Network (eSCL/WiFi)',wia:'USB (WIA)',twain:'USB (TWAIN)',naps2:'NAPS2'}[u.method]||u.method||'—';
  var share=(D.total>0)?((u.scans*100/D.total).toFixed(u.scans*100/D.total<1?2:1)+'%'):'—';
  var H=uHealth(u), hc=H.s>=80?'var(--ok)':(H.s>=60?'var(--warn)':'var(--bad)');
  var dm=u.daysMap||{};
  function since(n){var s=0,dt=new Date();for(var i=0;i<n;i++){var d=new Date(dt.getTime()-i*86400000).toISOString().slice(0,10);s+=parseInt(dm[d])||0;}return s;}
  var crashes=(D.crashes||[]).filter(function(c){return c.client&&u.id&&String(c.client)===String(u.id).slice(0,40);});
  var fbs=(D.feedback||[]).filter(function(f){return (f.client&&String(f.client)===String(u.id).slice(0,40))||(f.name&&u.name!=='—'&&f.name===u.name);});
  function kv(l,v){ return '<span>'+l+'</span><b>'+v+'</b>'; }
  // achievements (asli data se)
  var actDays=(u.active&&u.active.length)||0;
  var ACH=[['🥉','100 Scans',u.scans>=100],['🥈','500 Scans',u.scans>=500],['🥇','1000 Scans',u.scans>=1000],
    ['📅','30 Active Days',actDays>=30],['🔥','Aaj active',uToday(u)>0],['⚡','Power User',u.scans>=Math.max(50,(D.total||1)*0.1)]];
  // AI insights (rules — sirf asli data)
  var ins=[]; var w1=since(7), w2=since(14)-since(7);
  if(w2>0&&w1>w2*1.25) ins.push(['📈','Is hafte scanning <b>'+Math.round((w1-w2)*100/w2)+'% badhi</b> hai pichhle hafte se.']);
  if(w2>0&&w1<w2*0.6) ins.push(['📉','Is hafte scanning <b>'+Math.round((w2-w1)*100/Math.max(1,w2))+'% ghati</b> — dhyaan dein.']);
  if(u.version&&D.latestVersion&&String(u.version).trim()!==String(D.latestVersion).trim()) ins.push(['⚠','Purane version v'+esc(u.version)+' par hai (latest v'+D.latestVersion+') — 📨 Message se update yaad dilayein.']);
  var cr7=crashes.filter(function(c){return (Date.now()/1000-(c.t||0))<7*86400;}).length;
  if(cr7) ins.push(['💥','Pichhle 7 din me <b>'+cr7+' crash</b> — jaanch zaroori.']);
  var offd=u.last?Math.floor((Date.now()/1000-u.last)/86400):0;
  if(offd>=14) ins.push(['😴','<b>'+offd+' din</b> se offline — follow-up ka time.']);
  if(!ins.length) ins.push(['✅','Sab normal — koi khaas dhyaan nahi chahiye.']);
  var html=
   '<div style="display:flex;gap:14px;align-items:center;flex-wrap:wrap;margin-bottom:4px">'
    +uAvatar(u,true)
    +'<div style="flex:1;min-width:150px"><div style="font-size:20px;font-weight:800">'+esc(u.name!=='—'?u.name:'(bina naam)')+'</div>'
    +'<div style="margin-top:4px">'+uStatus(u)+'</div>'
    +'<div style="color:var(--mut);font-size:11px;margin-top:4px">'+(u.online?'Abhi online':'Last seen '+ago(u.last)+' pehle')+' · Joined '+(u.first?new Date(u.first*1000).toLocaleDateString():'—')+'</div></div>'
    +'<div class="hring" style="background:conic-gradient('+hc+' '+(H.s*3.6)+'deg,var(--line) 0)"><span style="background:var(--card);width:48px;height:48px;border-radius:50%;display:grid;place-items:center;color:'+hc+'">'+H.s+'%</span></div>'
   +'</div>'
   +'<div style="color:var(--mut);font-size:10px;margin-bottom:6px">Health: '+esc(H.why)+'</div>'
   +'<div class="kpis" style="grid-template-columns:repeat(auto-fill,minmax(110px,1fr));margin-top:8px">'
    +_kpi('','📄',fmt(u.scans),'Total scans')+_kpi('g','📅',fmt(since(1)),'Aaj')
    +_kpi('p','🗓',fmt(since(7)),'7 din')+_kpi('y','📆',fmt(since(30)),'30 din')
    +_kpi('o','🌍',share,'World share')+_kpi('','⚡',actDays,'Active din')
   +'</div>'
   +'<div class="profgrid">'
    +'<div class="pcard"><h4>👤 Overview</h4><div class="pkv">'
      +kv('User ID',esc(String(u.id).slice(0,18))+'…')+kv('Country',(u.country?flag(u.country)+' '+u.country:'—'))
      +((u.region||u.city)?kv('Region',esc([u.city,u.region].filter(Boolean).join(', '))):'')
      +kv('IP',esc(u.ip||'—'))+kv('Joined',u.first?new Date(u.first*1000).toLocaleString():'—')
      +kv('Last activity',u.last?new Date(u.last*1000).toLocaleString():'—')
      +(u.note?kv('📝 Note','<i>'+esc(u.note)+'</i>'):'')
    +'</div></div>'
    +'<div class="pcard"><h4>🖨 Device &amp; Scanner</h4><div class="pkv">'
      +kv('Scanner',esc(u.model||'—'))+kv('Connection',mth)
      +kv('App version',(u.version?'v'+u.version:'—')+(u.version&&D.latestVersion&&u.version!==D.latestVersion?' <span class="stb w">⚠ old</span>':' <span class="stb g">✓ latest</span>'))
      +kv('License','<span class="stb g">Free · lifetime</span>')
      +kv('Crashes',crashes.length?('<span style="color:var(--bad);font-weight:700">'+crashes.length+'</span>'):'0 🎉')
    +'</div></div>'
   +'</div>'
   +((u.feats&&u.feats.length)?'<div class="pcard" style="margin-top:11px"><h4>🧰 Feature usage</h4>'+u.feats.map(function(f){return '<span class="tag">'+esc(FEATLBL[f]||f)+'</span>';}).join(' ')+'</div>':'')
   +'<div class="pcard" style="margin-top:11px"><h4>🏆 Achievements</h4>'
     +ACH.map(function(a){return '<span class="ach'+(a[2]?'':' off')+'">'+a[0]+' '+a[1]+'</span>';}).join('')+'</div>'
   +'<div class="pcard" style="margin-top:11px"><h4>🤖 AI Insights</h4>'
     +ins.map(function(i){return '<div class="aii"><span>'+i[0]+'</span><div>'+i[1]+'</div></div>';}).join('')+'</div>'
   +(function(){
      var mx=1; for(var kk in dm){ if((parseInt(dm[kk])||0)>mx) mx=parseInt(dm[kk]); }
      var strip='',dt=new Date();
      for(var i=59;i>=0;i--){ var dd=new Date(dt.getTime()-i*86400000).toISOString().slice(0,10); var c=parseInt(dm[dd])||0;
        var a=c>0?(0.25+0.75*Math.min(1,c/mx)):0; var bg=c>0?('rgba(59,130,246,'+a.toFixed(2)+')'):'var(--line)';
        strip+='<span title="'+dd+': '+c+' scan" style="width:9px;height:9px;border-radius:2px;background:'+bg+'"></span>'; }
      var keys=Object.keys(dm).filter(function(k){return (parseInt(dm[k])||0)>0;}).sort().reverse().slice(0,90);
      var hist=keys.map(function(k){return '<tr><td style="color:var(--mut);padding:2px 0">'+k+'</td><td style="text-align:right;font-weight:600">'+fmt(dm[k])+' scan</td></tr>';}).join('');
      return '<div class="pcard" style="margin-top:11px"><h4>📊 Scan history (60-din heatmap + din-wise)</h4>'
        +'<div style="display:flex;flex-wrap:wrap;gap:2px;margin-bottom:8px">'+strip+'</div>'
        +(hist?'<div style="max-height:150px;overflow:auto"><table style="font-size:11px">'+hist+'</table></div>':'<div style="color:var(--mut);font-size:11px">— abhi record nahi —</div>')+'</div>';
    })()
   +((u.featc&&u.featc.length)?(function(){ var mx=u.featc[0][1]||1;
      return '<div class="pcard" style="margin-top:11px"><h4>⭐ Ye user kya SABSE ZYADA karta hai</h4>'
       +u.featc.map(function(f){ return '<div style="display:flex;align-items:center;gap:7px;margin:4px 0"><span style="width:120px;font-size:11px">'+esc(flbl(f[0]))+'</span><div class="hbar" style="flex:1;height:8px"><i style="width:'+Math.round(f[1]*100/mx)+'%;background:linear-gradient(90deg,var(--accent),var(--accent2))"></i></div><b style="font-size:11px;min-width:34px;text-align:right">'+fmt(f[1])+'</b></div>'; }).join('')+'</div>'; })():'')
   +((u.ev&&u.ev.length)?'<div class="pcard" style="margin-top:11px"><h4>🕒 Timeline — is user ki haal ki activity</h4><table style="font-size:11px">'
      +u.ev.map(function(e){ return '<tr><td>⚙️ '+esc(flbl(e[1]))+'</td><td style="text-align:right;color:var(--mut);white-space:nowrap">'+ago(e[0])+'</td></tr>'; }).join('')+'</table></div>':'')
   +(crashes.length?'<div class="pcard" style="margin-top:11px"><h4>💥 Crash history</h4><div style="max-height:120px;overflow:auto"><table style="font-size:11px">'
      +crashes.slice(0,15).map(function(c){return '<tr><td>'+esc((c.err||'').slice(0,50))+'</td><td style="text-align:right;color:var(--mut);white-space:nowrap">v'+(c.v||'?')+' · '+ago(c.t)+'</td></tr>';}).join('')+'</table></div></div>':'')
   +(fbs.length?'<div class="pcard" style="margin-top:11px"><h4>💬 Feedback history</h4>'
      +fbs.slice(0,6).map(function(f){return '<div class="aii"><span>'+('★'.repeat(f.rating||0)||'—')+'</span><div>'+esc((f.msg||'').slice(0,90))+(f.reply?' <span class="stb g">↩ replied</span>':' <span class="stb y">pending</span>')+'</div></div>';}).join('')+'</div>':'')
   +'<div class="pcard" style="margin-top:11px"><h4>✏️ Manage</h4>'
     +'<form method="post" style="margin-bottom:6px;display:flex;gap:6px"><input type="hidden" name="act" value="rename"><input type="hidden" name="id" value="'+esc(u.id)+'"><input name="name" value="'+esc(u.name!=='—'?u.name:'')+'" placeholder="Naam badlo" style="flex:1"><button class="btn">Rename</button></form>'
     +'<form method="post" style="margin-bottom:6px;display:flex;gap:6px"><input type="hidden" name="act" value="tag"><input type="hidden" name="id" value="'+esc(u.id)+'"><input name="tags" value="'+esc(u.tags)+'" placeholder="Tags: VIP,Hospital,Beta" style="flex:1"><button class="btn">Tag</button></form>'
     +'<form method="post" style="margin-bottom:6px;display:flex;gap:6px"><input type="hidden" name="act" value="note"><input type="hidden" name="id" value="'+esc(u.id)+'"><input name="note" value="'+esc(u.note)+'" placeholder="Admin note (scanner badla, call kiya…)" style="flex:1"><button class="btn">Note</button></form>'
     +'<form method="post" style="margin-bottom:6px;display:flex;gap:6px"><input type="hidden" name="act" value="umsg"><input type="hidden" name="id" value="'+esc(u.id)+'"><input name="msg" placeholder="Is user ko message (app me dikhega)" style="flex:1"><button class="btn">📨 Message</button></form>'
     +'<button class="btn gray" onclick="uCSV([JSON.parse(this.getAttribute(\'data-u\'))])" data-u="'+esc(JSON.stringify({name:u.name,online:u.online,model:u.model,scans:u.scans,daysMap:{},last:u.last,first:u.first,version:u.version,country:u.country,region:u.region,method:u.method,id:u.id}))+'">⬇ User report (CSV)</button>'
   +'</div>';
  document.getElementById('umbody').innerHTML=html;
  document.getElementById('umodal').style.display='flex';
}
function closeUser(){ document.getElementById('umodal').style.display='none'; }
document.getElementById('usearch').addEventListener('input',renderUsers);
renderUsers();
// ---- (v2.2) Feature champions (kaun kya sabse zyada karta hai) + activity feed ----
(function(){ var el=document.getElementById('uchamp'); if(!el)return;
  var tot={}, top={};
  (D.userList||[]).forEach(function(u){ (u.featc||[]).forEach(function(f){
    tot[f[0]]=(tot[f[0]]||0)+f[1];
    if(!top[f[0]]||f[1]>top[f[0]][1]) top[f[0]]=[u.name,f[1],u.id]; }); });
  var rows=Object.keys(tot).sort(function(a,b){return tot[b]-tot[a];}).slice(0,12);
  el.classList.remove('skel');
  el.innerHTML=rows.length?_tbl(rows.map(function(k){ var t2=top[k];
    return '<tr><td>⚙️ '+esc(flbl(k))+'</td><td><a href="#" style="color:var(--accent);text-decoration:none" onclick="openUser(\''+esc(t2[2])+'\');return false"><b>'+esc(t2[0])+'</b></a> <span style="color:var(--mut)">('+fmt(t2[1])+' baar)</span></td><td style="text-align:right;color:var(--mut)">kul '+fmt(tot[k])+'</td></tr>';}),['Feature','#1 user','Sab total'])
   :'<div style="color:var(--mut);font-size:12px">— jaise-jaise users features chalaayenge, yahan champions dikhenge —</div>';
  var fe=document.getElementById('uactfeed'); if(!fe)return; fe.classList.remove('skel');
  var E=(D.recentEvents||[]).slice(0,40);
  fe.innerHTML=E.length?_tbl(E.map(function(e){
    return '<tr><td><b>'+esc(e.u||'user')+'</b></td><td>'+esc(flbl(e.feat))+'</td><td style="text-align:right;color:var(--mut);white-space:nowrap">'+ago(e.t)+'</td></tr>';}),['User','Kaam','Kab'])
   :'<div style="color:var(--mut);font-size:12px">— abhi koi event record nahi (v159+ app se events aate hain) —</div>';
})();

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
  kpi('🕒',fmt(D.last24h),'Last 24h')+
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
function flbl(k){ return FEATLBL[k]||k; }
// impact tiles
document.getElementById('impact').innerHTML=
  kpi('📑',fmt(D.impactDocs),'Documents bane')+
  kpi('📄',fmt(D.impactPaper),'Paper digitize','g')+
  kpi('💾',(D.impactDataMB>=1024?(D.impactDataMB/1024).toFixed(1)+' GB':D.impactDataMB+' MB'),'Data bachaya','p')+
  kpi('🌳',D.impactTrees,'Ped bachaye (~)','g')+
  kpi('🕒',fmt(D.impactHours)+'h','Samay bachaya','o')+
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
  var start='overview'; try{ var s=localStorage.getItem('anpage'); if(s) start=s; }catch(e){} try{ var qp=new URLSearchParams(location.search).get('p'); if(qp) start=qp; }catch(e){}
  if(!document.querySelector('.tab[data-p="'+start+'"]')) start='overview';
  show(start);
})();

// auto-refresh 30s (agar koi modal/form khula nahi)
setInterval(function(){ if(document.getElementById('umodal').style.display==='flex')return; if(document.activeElement&&/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName))return; location.reload(); },30000);

// ================= Fluent shell: sidebar / bell / search / counters =================
function sbToggle(force){ var sb=document.getElementById('sb'),ov=document.getElementById('sbOverlay');
  var open=(force===undefined)?!sb.classList.contains('open'):force;
  sb.classList.toggle('open',open); ov.classList.toggle('on',open); }
function jumpTo(page,cardId){ sbToggle(false);
  var tabs=[].slice.call(document.querySelectorAll('.sb .tab[data-p]'));
  var t=tabs.filter(function(b){return b.getAttribute('data-p')===page;})[0];
  if(t) t.click();
  setTimeout(function(){ var c=document.getElementById(cardId);
    if(c){ c.scrollIntoView({behavior:'smooth',block:'start'}); c.style.boxShadow='0 0 0 3px rgba(59,130,246,.5)'; setTimeout(function(){c.style.boxShadow='';},1600); } },120); }
// ---- notification bell: naya feedback + crashes (unread = pichhli visit ke baad) ----
(function(){ var items=[];
  (D.feedback||[]).slice(0,6).forEach(function(f){ items.push({t:f.t||0,ic:'💬',txt:(f.name||'user')+': '+String(f.msg||'').slice(0,60)}); });
  (D.crashes||[]).slice(0,6).forEach(function(c){ items.push({t:c.t||0,ic:'💥','txt':'Crash v'+(c.v||'?')+': '+String(c.err||'').slice(0,60)}); });
  (D.failLogRecent||[]).slice(0,4).forEach(function(f){ items.push({t:f.t||0,ic:'🚨',txt:'Galat admin-password koshish ('+(f.ip||'?')+')'}); });
  var SYB=D.sys||{}; if(SYB.diskFreePct!==undefined&&SYB.diskFreePct<10) items.push({t:Date.now()/1000,ic:'💽',txt:'Server disk sirf '+SYB.diskFreePct+'% khaali!'});
  items.sort(function(a,b){return (b.t||0)-(a.t||0);}); items=items.slice(0,10);
  var seen=0; try{ seen=parseInt(localStorage.getItem('an_nseen')||'0'); }catch(e){}
  var unread=items.filter(function(i){return (i.t||0)>seen;}).length;
  var badge=document.getElementById('nbadge');
  if(unread>0&&badge){ badge.style.display='flex'; badge.textContent=unread>9?'9+':unread; }
  var nd=document.getElementById('ndrop');
  nd.innerHTML = items.length ? items.map(function(i){
      return '<div class="ni"><span>'+i.ic+'</span><div>'+i.txt.replace(/[&<>]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c];})+'<div class="t">'+ago(i.t)+' ago</div></div></div>';
    }).join('') : '<div class="ni" style="color:var(--mut)">— koi nayi notification nahi —</div>';
  window.bellToggle=function(){ var o=nd.classList.toggle('open');
    if(o){ try{ localStorage.setItem('an_nseen',String(Math.floor(Date.now()/1000))); }catch(e){}
      if(badge)badge.style.display='none'; } };
  document.addEventListener('click',function(e){ if(!document.getElementById('bellBtn').contains(e.target)) nd.classList.remove('open'); });
})();
// ---- global search: users page ke search se juda + Ctrl+K ----
(function(){ var g=document.getElementById('gsearch'), u=document.getElementById('usearch');
  if(!g)return;
  g.addEventListener('input',function(){ if(!u)return; u.value=g.value;
    if(g.value.length>=2){ var t=document.querySelector('.sb .tab[data-p=users]'); if(t&&!t.classList.contains('active'))t.click(); }
    u.dispatchEvent(new Event('input')); });
  document.addEventListener('keydown',function(e){ if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='k'){ e.preventDefault(); g.focus(); } });
})();
// ---- KPI counter animation (numbers count up on load) ----
(function(){ if(matchMedia('(prefers-reduced-motion: reduce)').matches)return;
  [].slice.call(document.querySelectorAll('.kpi .n')).forEach(function(el){
    var txt=el.textContent, num=parseFloat(txt.replace(/[^0-9.]/g,''));
    if(!isFinite(num)||num<=0)return;
    var pre='',suf=''; var m=txt.match(/^([^0-9]*)([\d,\.]+)(.*)$/); if(m){pre=m[1];suf=m[3];}
    var t0=null;
    function tick(ts){ if(!t0)t0=ts; var p=Math.min((ts-t0)/900,1); p=1-Math.pow(1-p,3);
      el.textContent=pre+Math.round(num*p).toLocaleString()+suf;
      if(p<1)requestAnimationFrame(tick); else el.textContent=txt; }
    requestAnimationFrame(tick); });
})();
// ================= PORTAL MODULES (naye pages — sab D data se) =================
function jumpToPage(p){ var t=document.querySelector('.sb .tab[data-p='+p+']'); if(t)t.click(); }
function _tbl(rows,heads){ return '<div style="max-height:420px;overflow:auto"><table><thead><tr>'+heads.map(function(h){return '<th>'+h+'</th>';}).join('')+'</tr></thead><tbody>'+rows.join('')+'</tbody></table></div>'; }
function _kpi(cls,ic,n,l){ return '<div class="kpi '+cls+'"><div class="ic">'+ic+'</div><div class="tx"><div class="n">'+n+'</div><div class="l">'+l+'</div></div></div>'; }
// ================= MISSION CONTROL (Live Monitoring) =================
// Sab kuch ASLI data se: userList / recentScans300 / crashes / feedback /
// hoursAll / daysMap / heat7x24 / usersByCountry / health. Koi fake metric nahi.
function lvToday(u){ return uToday(u); }
function lvActivity(u){
  // aakhri asli scan-event se activity nikaalo (recentScans me naam-match)
  var ev=(D.recentScans300||D.recentScans||[]).filter(function(x){return x.name&&u.name!=='—'&&x.name===u.name;})[0];
  if(u.online){
    if(ev&&(Date.now()/1000-(ev.t||0))<600) return '<span class="stb g">🖨 Scanning · '+ev.n+'p ('+ago(ev.t)+')</span>';
    return '<span class="stb g">🟢 Idle (online)</span>';
  }
  if(ev&&(Date.now()/1000-(ev.t||0))<86400) return '<span class="stb y">🖨 Last scan '+ago(ev.t)+'</span>';
  return '<span class="stb r">⚪ Offline</span>';
}
var _lf='all';
function lvRows(){
  var q=(document.getElementById('lvQ').value||'').toLowerCase();
  var sc=(document.getElementById('lvScn')||{}).value||'';
  return (D.userList||[]).filter(function(u){
    if(q&&(u.name+' '+(u.model||'')+' '+(u.country||'')+' '+(u.region||'')+' '+(u.ip||'')+' '+(u.version||'')).toLowerCase().indexOf(q)<0)return false;
    if(sc&&(u.model||'')!==sc)return false;
    if(_lf==='online'&&!u.online)return false;
    if(_lf==='offline'&&u.online)return false;
    if(_lf==='today'&&lvToday(u)<=0)return false;
    if(_lf==='oldver'&&!(u.version&&D.latestVersion&&String(u.version).trim()!==String(D.latestVersion).trim()))return false;
    return true;
  }).sort(function(a,b){ return (b.online-a.online)||((b.last||0)-(a.last||0)); });
}
function lvCSV(which){
  var rows=lvRows(); if(which==='online')rows=rows.filter(function(u){return u.online;});
  uCSV(rows);
}
function lvRender(){
  var U=D.userList||[]; var on=U.filter(function(u){return u.online&&!u.blocked;});
  var scanners={}; U.forEach(function(u){ if(u.model)scanners[u.model]=1; });
  var activeScn={}; on.forEach(function(u){ if(u.model)activeScn[u.model]=1; });
  // header strip
  document.getElementById('lvResp').textContent=(D.respMs||0)+' ms';
  document.getElementById('lvOnN').textContent=on.length;
  document.getElementById('lvDevN').textContent=Object.keys(activeScn).length;
  // scanner filter options
  (function(){ var sel=document.getElementById('lvScn'); if(sel&&sel.options.length<=1)
    Object.keys(scanners).sort().forEach(function(m){ var o=document.createElement('option');o.value=m;o.textContent='🖨 '+m;sel.appendChild(o); }); })();
  // ---- KPIs (sab asli) ----
  var today=D.today||0, yest=D.yesterday||0;
  var trend=yest>0?Math.round((today-yest)*100/yest):null;
  var crashesToday=(D.crashes||[]).filter(function(c){return (Date.now()/1000-(c.t||0))<86400;}).length;
  var avgHealth=U.length?Math.round(U.reduce(function(a,u){return a+uHealth(u).s;},0)/U.length):100;
  var H=D.health||{};
  var bakAge=D.lastBackup?Math.floor((Date.now()/1000-D.lastBackup)/3600):null;
  document.getElementById('liveKpis').innerHTML=
    _kpi('','👥',fmt(U.length),'Registered users')
    +_kpi('g','🟢',on.length,'Online abhi')
    +_kpi('r','🔴',U.length-on.length,'Offline')
    +_kpi('p','🖨',Object.keys(scanners).length,'Scanner models')
    +_kpi('g','📄',fmt(today)+(trend!==null?' <small style="font-size:9px;color:'+(trend>=0?'var(--ok)':'var(--bad)')+'">'+(trend>=0?'▲':'▼')+Math.abs(trend)+'%</small>':''),"Aaj ke scans (pages)")
    +_kpi('','📊',(D.avgPages||0),'Avg pages / scan')
    +_kpi(avgHealth>=80?'g':(avgHealth>=60?'y':'r'),'⭐',avgHealth+'%','Fleet health score')
    +_kpi('','💾',(D.fileKB||0)+' KB','JSON storage')
    +_kpi(bakAge===null?'y':(bakAge<48?'g':'r'),'📦',bakAge===null?'—':(bakAge<1?'<1h':bakAge+'h'),'Last backup')
    +_kpi(D.bcMsg?'g':'','📢',D.bcMsg?'Active':'—','Broadcast')
    +_kpi('p','💬',(D.feedback||[]).length,'Feedback stored')
    +_kpi(crashesToday?'r':'g','💥',crashesToday,'Crashes — 24h');
  // ---- version coverage ----
  (function(){ var lat=String(D.latestVersion||'').trim(); var out=0,onl=0;
    U.forEach(function(u){ if(u.version){ if(String(u.version).trim()===lat)onl++; else out++; } });
    var tot=onl+out, pct=tot?Math.round(onl*100/tot):100;
    document.getElementById('lvVer').innerHTML=
     '<div style="font-size:26px;font-weight:900;color:'+(pct>=70?'var(--ok)':'var(--warn)')+'">'+pct+'%</div>'
     +'<div style="color:var(--mut);font-size:10.5px;margin-bottom:6px">users latest v'+esc(lat||'?')+' par hain</div>'
     +'<div class="hbar" style="width:100%;height:9px"><i style="width:'+pct+'%;background:linear-gradient(90deg,var(--accent),var(--accent2))"></i></div>'
     +'<div style="display:flex;justify-content:space-between;font-size:10px;color:var(--mut);margin-top:5px"><span>✓ '+onl+' latest</span><span>⚠ '+out+' purane</span></div>'
     +(out?'<button class="btn" style="margin-top:8px;width:100%" onclick="jumpTo(\'overview\',\'cardBroadcast\')">📣 Update broadcast bhejo</button>':'');
  })();
  // ---- Alert Center (rules — sirf asli haalat par) ----
  (function(){ var A=[];
    var lat=String(D.latestVersion||'').trim();
    var oldv=U.filter(function(u){return u.version&&String(u.version).trim()!==lat;}).length;
    if(crashesToday) A.push(['crit','💥',crashesToday+' crash pichhle 24h me — Crash Center dekho']);
    if(bakAge!==null&&bakAge>=48) A.push(['crit','📦','Backup '+bakAge+' ghante purana — naya backup lo']);
    if(oldv) A.push(['warn','⚠',oldv+' user purane version par — update broadcast bhejo']);
    var inact=U.filter(function(u){return u.last&&(Date.now()/1000-u.last)>14*86400;}).length;
    if(inact) A.push(['warn','😴',inact+' user 14+ din se offline']);
    if((D.fileKB||0)>2048) A.push(['warn','💾','stats.json '+Math.round(D.fileKB/1024)+' MB ka ho gaya — purge/cleanup socho']);
    // (v2) naye REAL rules: disk, crash-spike, failed logins, CPU load
    var SY=D.sys||{};
    if(SY.diskFreePct!==undefined&&SY.diskFreePct<10) A.push(['crit','💽','Server disk sirf '+SY.diskFreePct+'% khaali ('+SY.diskFreeGB+' GB) — jagah banao']);
    if(SY.loadPct!==undefined&&SY.loadPct>90) A.push(['warn','💻','Server CPU '+SY.loadPct+'% par hai (load '+SY.load+', '+SY.cores+' cores)']);
    (function(){ var now=Date.now()/1000, c24=0, c7=0;
      (D.crashes||[]).forEach(function(c){ var a=now-(c.t||0); if(a<86400)c24++; if(a<7*86400)c7++; });
      var avg=(c7-c24)/6;
      if(c24>=3&&c24>2*Math.max(0.5,avg)) A.push(['crit','📈','CRASH SPIKE: aaj '+c24+' crash (7-din avg ~'+avg.toFixed(1)+'/din) — nayi release jaancho']); })();
    (function(){ var now=Date.now()/1000, f=(D.failLogRecent||[]).filter(function(x){return now-(x.t||0)<3600;}).length;
      if(f>=3) A.push(['crit','🚨','Suraksha: pichhle ghante me '+f+' galat-password koshish — Activity Logs dekho']); })();
    if(H.lastRecovery) A.push(['info','♻️','Storage auto-recovery chali thi ('+new Date(H.lastRecovery*1000).toLocaleString()+') — sab theek hai']);
    var blocked=U.filter(function(u){return u.blocked;}).length;
    if(blocked) A.push(['info','🚫',blocked+' user blocked hain']);
    if(!A.length) A.push(['info','✅','Koi alert nahi — sab systems normal 🎉']);
    window._lvA=A; lvAlertsRender('all');
  })();
  // ---- AI Insights ----
  (function(){ var I=[];
    if(trend!==null&&Math.abs(trend)>=15) I.push([trend>0?'📈':'📉','Aaj scanning kal se <b>'+Math.abs(trend)+'% '+(trend>0?'zyada':'kam')+'</b> hai.']);
    var lat=String(D.latestVersion||'').trim(); var tot=0,onl=0;
    U.forEach(function(u){ if(u.version){tot++; if(String(u.version).trim()===lat)onl++;} });
    if(tot) I.push(['🔢','<b>'+Math.round(onl*100/tot)+'%</b> users latest version par hain.']);
    var topScn='',mx=0,cnt={}; U.forEach(function(u){ if(u.model){cnt[u.model]=(cnt[u.model]||0)+(u.scans||0); if(cnt[u.model]>mx){mx=cnt[u.model];topScn=u.model;}} });
    if(topScn) I.push(['🖨','Sabse zyada kaam <b>'+esc(topScn)+'</b> par ho raha hai ('+fmt(mx)+' scans).']);
    var hh=D.hoursAll||{}; var ph=-1,pm=0; Object.keys(hh).forEach(function(k){ var h=parseInt(k.slice(-2)); var v=parseInt(hh[k])||0; if(v>pm){pm=v;ph=h;} });
    if(ph>=0) I.push(['🕒','Peak scanning time: <b>'+ph+':00–'+(ph+1)+':00</b> ke aas-paas.']);
    var inact=U.filter(function(u){return u.last&&(Date.now()/1000-u.last)>7*86400;}).length;
    if(inact) I.push(['😴','<b>'+inact+' user</b> 7+ din se nahi aaye — wapas laane ka message bhejo.']);
    if(bakAge!==null&&bakAge>=24) I.push(['📦','Aakhri backup <b>'+bakAge+'h</b> pehle — aaj ka backup le lo.']);
    if(!I.length) I.push(['✅','Sab stable — koi khaas pattern nahi.']);
    document.getElementById('lvAI').innerHTML=I.map(function(i){return '<div class="aii"><span>'+i[0]+'</span><div>'+i[1]+'</div></div>';}).join('');
  })();
  // ---- live grid ----
  (function(){ var rows=lvRows();
    var h='<div style="max-height:420px;overflow:auto"><table><thead><tr><th>User</th><th>Activity</th><th>Scanner</th><th>Aaj</th><th>Ver</th><th>Location</th><th>IP</th><th>Heartbeat</th><th>Health</th><th class="no-print">⚡</th></tr></thead><tbody>'
     +(rows.length?rows.map(function(u){ var hh=uHealth(u).s, hc=hh>=80?'var(--ok)':(hh>=60?'var(--warn)':'var(--bad)');
       return '<tr><td style="cursor:pointer;min-width:130px" onclick="showUser(lvFind(\''+esc(u.id)+'\'))"><div style="display:flex;gap:7px;align-items:center">'+uAvatar(u)+'<b>'+esc(u.name)+'</b></div></td>'
        +'<td>'+lvActivity(u)+'</td><td>'+esc((u.model||'—').slice(0,20))+'</td>'
        +'<td>'+(lvToday(u)?'<b style="color:var(--ok)">'+lvToday(u)+'p</b>':'<span style="color:var(--mut)">0</span>')+'</td>'
        +'<td>'+(u.version?('v'+u.version+(D.latestVersion&&String(u.version).trim()!==String(D.latestVersion).trim()?' <span class="stb w">old</span>':'')):'—')+'</td>'
        +'<td>'+(u.country?flag(u.country)+' '+u.country:'—')+(u.region?' · <span style="color:var(--mut)">'+esc(u.region)+'</span>':'')+'</td>'
        +'<td style="font-size:9.5px;color:var(--mut)">'+esc(u.ip||'—')+'</td>'
        +'<td>'+(u.online?'<span class="stb g">live</span>':ago(u.last))+'</td>'
        +'<td><span class="hbar"><i style="width:'+hh+'%;background:'+hc+'"></i></span></td>'
        +'<td class="no-print"><span class="uact"><button title="Message" onclick="event.stopPropagation();showUser(lvFind(\''+esc(u.id)+'\'));setTimeout(function(){var f=document.querySelector(\'#umbody input[name=msg]\');if(f)f.focus();},150)">📨</button></span></td></tr>';
     }).join(''):'<tr><td colspan="10" style="text-align:center;color:var(--mut);padding:18px">— koi user match nahi hua —</td></tr>')
     +'</tbody></table></div>';
    document.getElementById('lvGrid').innerHTML=h;
  })();
  // ---- charts (Chart.js pehle se loaded; na ho to skip) ----
  if(window.Chart){
    try{
      var hrs=[],vals=[]; for(var i=23;i>=0;i--){ var d=new Date(Date.now()-i*3600000);
        var k=d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0')+'-'+String(d.getHours()).padStart(2,'0');
        hrs.push(d.getHours()+':00'); vals.push(parseInt((D.hoursAll||{})[k])||0); }
      new Chart(document.getElementById('lv24'),{type:'bar',data:{labels:hrs,datasets:[{data:vals,backgroundColor:'rgba(59,130,246,.65)',borderRadius:4}]},options:{plugins:{legend:{display:false}},scales:{x:{ticks:{maxTicksLimit:8,color:'#94A3B8'},grid:{display:false}},y:{ticks:{color:'#94A3B8'},grid:{color:'rgba(148,163,184,.12)'}}}}});
      var days=[],dv=[]; for(var j=13;j>=0;j--){ var dd=new Date(Date.now()-j*86400000).toISOString().slice(0,10);
        days.push(dd.slice(5)); dv.push(parseInt((D.daysMap||{})[dd])||0); }
      new Chart(document.getElementById('lv14'),{type:'line',data:{labels:days,datasets:[{data:dv,fill:true,tension:.35,borderColor:'#3B82F6',backgroundColor:'rgba(59,130,246,.18)',pointRadius:2}]},options:{plugins:{legend:{display:false}},scales:{x:{ticks:{color:'#94A3B8'},grid:{display:false}},y:{ticks:{color:'#94A3B8'},grid:{color:'rgba(148,163,184,.12)'}}}}});
    }catch(e){}
  }
  // ---- heatmap 7x24 ----
  (function(){ var hm=D.heat7x24; var el=document.getElementById('lvHeat'); if(!el)return;
    if(!hm||!hm.length){ el.innerHTML='<div style="color:var(--mut);font-size:11px">— data nahi —</div>'; return; }
    var mx=1; hm.forEach(function(r){ r.forEach(function(v){ if(v>mx)mx=v; }); });
    var dn=['Som','Man','Bud','Gur','Shu','Sha','Rav'];
    var h='<div style="display:grid;grid-template-columns:34px repeat(24,1fr);gap:2px;font-size:8px">';
    h+='<span></span>'; for(var c=0;c<24;c++) h+='<span style="color:var(--mut);text-align:center">'+(c%4===0?c:'')+'</span>';
    hm.forEach(function(row,ri){ h+='<span style="color:var(--mut);line-height:14px">'+dn[ri]+'</span>';
      row.forEach(function(v,ci){ var a=v>0?(0.2+0.8*Math.min(1,v/mx)):0;
        h+='<span title="'+dn[ri]+' '+ci+':00 — '+v+' scan" style="height:14px;border-radius:3px;background:'+(v>0?'rgba(59,130,246,'+a.toFixed(2)+')':'var(--line)')+'"></span>'; }); });
    el.innerHTML=h+'</div>';
  })();
  // ---- country distribution ----
  (function(){ var el=document.getElementById('lvGeo'); if(!el)return; var cc=D.usersByCountry||[];
    if(!cc.length){ el.innerHTML='<div style="color:var(--mut);font-size:11px">— data nahi —</div>'; return; }
    var mx=cc[0]&&cc[0][1]||1;
    el.innerHTML=cc.slice(0,10).map(function(r){ var onc=(D.userList||[]).filter(function(u){return u.country===r[0]&&u.online;}).length;
      return '<div style="display:flex;gap:8px;align-items:center;margin-bottom:5px;font-size:11px">'
       +'<span style="width:52px">'+flag(r[0])+' '+esc(r[0])+'</span>'
       +'<span class="bar" style="flex:0 0 '+Math.max(6,Math.round(r[1]*100/mx))+'%;height:10px"></span>'
       +'<b>'+r[1]+'</b>'+(onc?' <span class="stb g">'+onc+' online</span>':'')+'</div>'; }).join('');
  })();
  // ---- timeline (asli events merge) ----
  (function(){ var el=document.getElementById('lvTime'); if(!el)return; var T=[];
    (D.recentScans300||D.recentScans||[]).slice(0,25).forEach(function(x){ T.push({t:x.t,ic:'🖨',txt:'<b>'+esc(x.name||'user')+'</b> ne <b>'+x.n+'</b> page scan kiye'+(x.cc?' ('+flag(x.cc)+')':'')}); });
    (D.recentEvents||[]).slice(0,25).forEach(function(e){ T.push({t:e.t,ic:'⚙️',txt:'<b>'+esc(e.u||'user')+'</b> ne <b>'+esc(flbl(e.feat))+'</b> chalaya'}); });
    (D.crashes||[]).slice(0,6).forEach(function(c){ T.push({t:c.t,ic:'💥',txt:'Crash v'+esc(c.v||'?')+' — '+esc((c.err||'').slice(0,40))}); });
    (D.feedback||[]).slice(0,6).forEach(function(f){ T.push({t:f.t,ic:'💬',txt:'<b>'+esc(f.name||'user')+'</b> ka feedback '+('★'.repeat(f.rating||0))}); });
    (D.auditLog||[]).slice(0,6).forEach(function(a){ T.push({t:a.t,ic:'🛠',txt:'Admin: '+esc(a.act)+(a.det?' — '+esc(a.det):'')}); });
    (D.adminLoginsFull||[]).slice(0,3).forEach(function(l){ if(l.t)T.push({t:l.t,ic:'🔐',txt:'Admin login ('+esc(l.ip||'?')+')'}); });
    T.sort(function(a,b){return (b.t||0)-(a.t||0);});
    el.innerHTML=T.length?T.slice(0,30).map(function(e){
      return '<div class="aii"><span>'+e.ic+'</span><div>'+e.txt+'<div style="color:var(--mut);font-size:9px">'+(e.t?new Date(e.t*1000).toLocaleTimeString():'')+' · '+ago(e.t)+' pehle</div></div></div>'; }).join('')
      :'<div style="color:var(--mut)">— abhi koi activity nahi —</div>';
  })();
}
function lvFind(id){ return (D.userList||[]).filter(function(u){return String(u.id)===id;})[0]; }
function lvAlertsRender(sev){
  var A=window._lvA||[]; if(sev!=='all')A=A.filter(function(a){return a[0]===sev;});
  var C={crit:'var(--bad)',warn:'var(--warn)',info:'var(--accent2)'};
  document.getElementById('lvAlerts').innerHTML=A.length?A.map(function(a){
    return '<div class="aii"><span style="color:'+C[a[0]]+'">'+a[1]+'</span><div>'+a[2]+' <span class="stb '+(a[0]==='crit'?'b':(a[0]==='warn'?'w':'g'))+'">'+a[0]+'</span></div></div>'; }).join('')
    :'<div style="color:var(--mut);font-size:11px">— is severity ka koi alert nahi —</div>';
}
(function(){ var f=document.getElementById('lvAlF'); if(f) [].forEach.call(f.querySelectorAll('.ufc'),function(b){
  b.onclick=function(){ [].forEach.call(f.querySelectorAll('.ufc'),function(x){x.classList.remove('on');}); b.classList.add('on'); lvAlertsRender(b.getAttribute('data-s')); }; }); })();
(function(){ var q=document.getElementById('lvQ'); if(q){ q.addEventListener('input',function(){ lvRender(); });
  var s=document.getElementById('lvScn'); if(s)s.addEventListener('change',lvRender);
  [].forEach.call(document.querySelectorAll('[data-lf]'),function(b){ b.onclick=function(){
    [].forEach.call(document.querySelectorAll('[data-lf]'),function(x){x.classList.remove('on');}); b.classList.add('on');
    _lf=b.getAttribute('data-lf'); lvRender(); }; }); } })();
// refresh countdown (page 30s par khud reload hoti hai)
(function(){ var n=30,el=document.getElementById('lvCount'); if(el) setInterval(function(){ n--; if(n<0)n=0; el.textContent=n; },1000); })();
try{ lvRender(); }catch(e){ if(window.jsonLog){} }
// ---- DEVICES (scanner-wise) ----
(function(){ var U=D.userList||[]; var g={};
  U.forEach(function(u){ var k=(u.model||'').trim()||'(Scanner naam nahi mila)';
    if(!g[k])g[k]={n:0,scans:0,last:0,vers:{},on:0,method:{}};
    g[k].n++; g[k].scans+=u.scans||0; g[k].last=Math.max(g[k].last,u.last||0);
    if(u.online)g[k].on++; if(u.version)g[k].vers[u.version]=1; if(u.method)g[k].method[u.method]=1; });
  var keys=Object.keys(g).sort(function(a,b){return g[b].scans-g[a].scans;});
  var kp=document.getElementById('hwKpis'); if(!kp)return;
  kp.innerHTML=_kpi('','🖨',keys.length,'Scanner models')+_kpi('g','🟢',keys.reduce(function(a,k){return a+g[k].on;},0),'Online devices')+_kpi('p','📄',fmt(keys.reduce(function(a,k){return a+g[k].scans;},0)),'Total scans');
  document.getElementById('hwCards').innerHTML = keys.length? keys.map(function(k){ var d=g[k];
    return '<div class="card" style="margin-bottom:0"><h3><span class="em">🖨</span> '+esc(k)+'</h3>'
      +'<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:8px;font-size:11px">'
      +'<div>👥 PCs<br><b>'+d.n+'</b></div><div>🟢 Online<br><b>'+d.on+'</b></div>'
      +'<div>📄 Scans<br><b>'+fmt(d.scans)+'</b></div><div>🕘 Last online<br><b>'+ago(d.last)+'</b></div>'
      +'<div>🔢 Versions<br><b>'+esc(Object.keys(d.vers).sort().map(function(v){return 'v'+v;}).join(', ')||'—')+'</b></div>'
      +'<div>🔌 Method<br><b>'+esc(Object.keys(d.method).join(', ')||'—')+'</b></div>'
      +'</div></div>'; }).join('') : '<div class="card" style="color:var(--mut)">— data nahi —</div>';
})();
// ---- SCANS grid: search + range + sort + pagination + CSV ----
var _sg={page:1,per:25,sort:'t',dir:-1,open:-1};
function _sgData(){ var rows=(D.recentScans300||D.recentScans||[]).slice();
  var q=(document.getElementById('sgQ').value||'').toLowerCase();
  var r=parseInt(document.getElementById('sgRange').value||'0');
  var fd=(document.getElementById('sgDpi')||{}).value||'';
  var fc=(document.getElementById('sgCol')||{}).value||'';
  var fs=(document.getElementById('sgSrc')||{}).value||'';
  var cut=r? (Math.floor(Date.now()/1000)-r*86400):0;
  rows=rows.filter(function(x){ if(cut&&(x.t||0)<cut)return false;
    if(fd&&String(x.dpi||'')!==fd)return false;
    if(fc&&String(x.col||'')!==fc)return false;
    if(fs&&String(x.src||'')!==fs)return false;
    if(q&&((x.name||'')+' '+(x.cc||'')+' '+(x.prof||'')+' '+(x.sm||'')).toLowerCase().indexOf(q)<0)return false; return true;});
  rows.sort(function(a,b){ var k=_sg.sort,va=a[k]||0,vb=b[k]||0;
    if(typeof va==='string'){va=va.toLowerCase();vb=(vb||'').toLowerCase();}
    return (va<vb?-1:va>vb?1:0)*_sg.dir;});
  return rows; }
function sgKpis(rows){ var el=document.getElementById('sgKpis'); if(!el)return;
  var today=new Date().toDateString(), tn=0, pg=0, colr=0, durS=0, durN=0;
  rows.forEach(function(x){ pg+=(x.n||0);
    if(new Date((x.t||0)*1000).toDateString()===today)tn++;
    if(String(x.col||'').toLowerCase().indexOf('colour')>-1)colr++;
    if(x.dur){durS+=x.dur;durN++;} });
  el.innerHTML=_kpi('','📄',fmt(rows.length),'Scans (filter ke baad)')
    +_kpi('g','🗓',tn,'Aaj ke scans')
    +_kpi('p','🧾',fmt(pg),'Kul pages')
    +_kpi('','📊',rows.length?(pg/rows.length).toFixed(1):'0','Avg pages/scan')
    +_kpi('y','🎨',rows.length?Math.round(100*colr/rows.length)+'%':'—','Colour scans')
    +(durN?_kpi('','⏱',Math.round(durS/durN)+'s','Avg samay/scan'):''); }
function sgRender(){ var el=document.getElementById('sgGrid'); if(!el)return;
  var rows=_sgData(); sgKpis(rows); var tp=Math.max(1,Math.ceil(rows.length/_sg.per));
  if(_sg.page>tp)_sg.page=tp;
  var pg=rows.slice((_sg.page-1)*_sg.per,_sg.page*_sg.per);
  function th(label,key){ var a=_sg.sort===key?(_sg.dir<0?' ↓':' ↑'):''; return '<th onclick="_sg.sort=\''+key+'\';_sg.dir=(_sg.sort===\''+key+'\'?-_sg.dir:-1);sgRender()">'+label+a+'</th>'; }
  var body='';
  if(pg.length){ pg.forEach(function(x,i){ var d=new Date((x.t||0)*1000); var gi=(_sg.page-1)*_sg.per+i;
      body+='<tr style="cursor:pointer" onclick="_sg.open=(_sg.open==='+gi+'? -1:'+gi+');sgRender()">'
        +'<td style="white-space:nowrap">'+d.toLocaleDateString()+' '+d.toLocaleTimeString()+'</td>'
        +'<td><b>'+esc(x.name||'—')+'</b></td><td>'+flag(x.cc)+' '+esc(x.cc||'—')+'</td>'
        +'<td><span class="tag">'+(x.n||0)+' pg</span></td>'
        +'<td>'+esc(x.dpi||'—')+'</td><td>'+esc(x.col||'—')+'</td>'
        +'<td>'+(x.src==='glass'?'🪟 Glass':(x.src?'📥 Feeder':'—'))+'</td>'
        +'<td>'+(x.dur?x.dur+'s':'—')+'</td></tr>';
      if(_sg.open===gi){ body+='<tr><td colspan="8" style="background:rgba(99,102,241,.06);font-size:11px;padding:9px 14px">'
        +'🖨 Scanner: <b>'+esc(x.sm||'—')+'</b> &nbsp;·&nbsp; 🗂 Profile: <b>'+esc(x.prof||'—')+'</b>'
        +' &nbsp;·&nbsp; 📐 Page size: <b>'+esc(x.sz||'—')+'</b> &nbsp;·&nbsp; 🔢 App: <b>'+(x.v?'v'+esc(x.v):'—')+'</b>'
        +' &nbsp;·&nbsp; ⏱ Samay: <b>'+(x.dur?x.dur+' sec':'—')+'</b> &nbsp;·&nbsp; 🌍 '+flag(x.cc)+' '+esc(x.cc||'—')
        +(x.rot?' &nbsp;·&nbsp; 🔁 Rotate: <b>'+esc(x.rot)+'</b>':'')
        +(x.tess===0?' &nbsp;·&nbsp; <b style="color:var(--warn)">⚠ Tesseract nahi</b>':'')
        +'</td></tr>'; }
    }); }
  else body='<tr><td colspan="8" style="color:var(--mut);text-align:center;padding:20px">— koi scan nahi mila —</td></tr>';
  el.innerHTML='<div style="max-height:440px;overflow:auto"><table><thead><tr>'+th('Date/Time','t')+th('User','name')+th('Country','cc')+th('Pages','n')+th('DPI','dpi')+th('Rang','col')+th('Srot','src')+th('Samay','dur')+'</tr></thead><tbody>'+body+'</tbody></table></div>';
  var pag=document.getElementById('sgPag');
  pag.innerHTML='<button class="btn gray" '+(_sg.page<=1?'disabled':'')+' onclick="_sg.page--;sgRender()">←</button>'
    +'<span style="color:var(--mut)">Page '+_sg.page+' / '+tp+' · '+fmt(rows.length)+' scans</span>'
    +'<button class="btn gray" '+(_sg.page>=tp?'disabled':'')+' onclick="_sg.page++;sgRender()">→</button>'; }
function sgCSV(){ var rows=_sgData();
  var csv='date,time,user,country,pages,dpi,colour,source,seconds,profile,scanner,app\n'+rows.map(function(x){ var d=new Date((x.t||0)*1000);
    return d.toLocaleDateString()+','+d.toLocaleTimeString()+',"'+String(x.name||'').replace(/"/g,'""')+'",'+(x.cc||'')+','+(x.n||0)+','+(x.dpi||'')+','+(x.col||'')+','+(x.src||'')+','+(x.dur||'')+',"'+String(x.prof||'').replace(/"/g,'""')+'","'+String(x.sm||'').replace(/"/g,'""')+'",'+(x.v||'');}).join('\n');
  var a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'})); a.download='apnescan-scans.csv'; a.click();
  showToast('⬇ Scans CSV download ho gayi','ok'); }
(function(){ var q=document.getElementById('sgQ'); if(!q)return;
  // DPI / Rang ke filter-option asli data se bhar do
  var ds={},cs={}; (D.recentScans300||[]).forEach(function(x){ if(x.dpi)ds[x.dpi]=1; if(x.col)cs[x.col]=1; });
  var sd=document.getElementById('sgDpi'); Object.keys(ds).sort(function(a,b){return a-b;}).forEach(function(k){ var o=document.createElement('option'); o.value=k; o.textContent=k+' dpi'; sd.appendChild(o); });
  var sc=document.getElementById('sgCol'); Object.keys(cs).sort().forEach(function(k){ var o=document.createElement('option'); o.value=k; o.textContent=k; sc.appendChild(o); });
  q.addEventListener('input',function(){_sg.page=1;_sg.open=-1;sgRender();});
  ['sgRange','sgDpi','sgCol','sgSrc'].forEach(function(id){ var e=document.getElementById(id); if(e)e.addEventListener('change',function(){_sg.page=1;_sg.open=-1;sgRender();}); });
  sgRender(); })();
// ---- (v3) DIN-BHAR KI POORI FEED — roz ki event-file se (har click) ----
var _df={rows:[],date:''};
function dfLoad(){ var dt=(document.getElementById('dfDate')||{}).value||'';
  fetch('?admin=1&api=feed'+(dt?'&date='+dt:''),{credentials:'same-origin'})
   .then(function(r){return r.json();}).then(function(j){
    if(!j||!j.ok)return; _df.rows=j.events||[]; _df.date=j.date;
    var sel=document.getElementById('dfDate');
    if(sel&&!sel.options.length){ (j.dates&&j.dates.length?j.dates:[j.date]).forEach(function(dd){
      var o=document.createElement('option'); o.value=dd; o.textContent=dd; sel.appendChild(o); }); sel.value=j.date; }
    var us={}; _df.rows.forEach(function(e){ if(e.u)us[e.u]=1; });
    var su=document.getElementById('dfUser');
    if(su){ var cur=su.value; su.innerHTML='<option value="">👥 Sab users</option>';
      Object.keys(us).sort().forEach(function(k){ var o=document.createElement('option'); o.value=k; o.textContent=k; su.appendChild(o); });
      su.value=cur; }
    dfRender(); }).catch(function(){});
}
function dfLbl(e){ var n=String(e||'');
  if(n.indexOf('btn:')===0)return '🔘 '+n.slice(4)+' (toolbar)';
  if(n.indexOf('menu:')===0)return '📋 '+n.slice(5)+' (menu)';
  if(n.indexOf('nav:')===0)return '🧭 '+n.slice(4)+' (sidebar)';
  if(n.indexOf('card:')===0)return '🃏 '+n.slice(5)+' (dashboard)';
  if(n==='app:start')return 'App KHOLI';
  if(n==='app:close')return 'App BAND ki';
  return (window.flbl?flbl(n):n); }
function dfRender(){ var el=document.getElementById('dfList'); if(!el)return;
  var fu=(document.getElementById('dfUser')||{}).value||'';
  var q=((document.getElementById('dfQ')||{}).value||'').toLowerCase();
  var rows=_df.rows.filter(function(e){ if(fu&&e.u!==fu)return false;
    if(q&&(String(e.e)+' '+(e.u||'')).toLowerCase().indexOf(q)<0)return false; return true;});
  rows.sort(function(a,b){return (a.t||0)-(b.t||0);});
  var out=[],last=0;
  rows.forEach(function(e){
    if(fu&&out.length&&((e.t-last>1800)||e.e==='app:start'))
      out.push('<div style="color:var(--accent2);font-size:10px;margin:8px 0 4px;border-top:1px dashed rgba(148,163,184,.3);padding-top:6px">— nayi session shuru —</div>');
    last=e.t;
    out.push('<div class="aii"><span>'+(e.e==='app:start'?'🟢':(e.e==='app:close'?'🔴':'•'))+'</span><div>'
      +(fu?'':'<b>'+esc(e.u||e.c||'user')+'</b> — ')+esc(dfLbl(e.e))
      +'<div style="color:var(--mut);font-size:9px">'+new Date((e.t||0)*1000).toLocaleTimeString()+'</div></div></div>');
  });
  var kp=document.getElementById('dfKpi');
  if(kp)kp.innerHTML='<b>'+rows.length+'</b> events · '+esc(_df.date||'')+(fu?' · <b>'+esc(fu)+'</b>':'');
  el.innerHTML=out.length?out.join(''):'<div style="color:var(--mut)">— is din ki koi event nahi (users ke paas app v183+ hone par aayegi) —</div>';
}
(function(){ var e1=document.getElementById('dfUser'); if(e1)e1.addEventListener('change',dfRender);
  var e2=document.getElementById('dfQ'); if(e2)e2.addEventListener('input',dfRender);
  var e3=document.getElementById('dfDate'); if(e3)e3.addEventListener('change',dfLoad);
  if(document.getElementById('dfList')) dfLoad(); })();
// ---- (v3.1) USER ACTIVITY — poora alag page: summary + charts + timeline ----
var _ua={rows:[],date:''};
function uaCat(e){ var n=String(e||'');
  if(n==='app:start'||n==='app:close')return 'app';
  if(n.indexOf('btn:')===0)return 'btn';
  if(n.indexOf('menu:')===0)return 'menu';
  if(n.indexOf('nav:')===0||n.indexOf('card:')===0)return 'nav';
  return 'feat'; }
// ---- (v3.3) BHAAG 1: Poori timeline — sab din, naya pehle, 100/page ----
var _tl={page:0,total:0};
function tlGo(p){ if(p<0)p=0; _tl.page=p;
  var q=((document.getElementById('tlQ')||{}).value||'').trim();
  fetch('?admin=1&api=tl&off='+(p*100)+'&lim=100'+(q?'&q='+encodeURIComponent(q):''),{credentials:'same-origin'})
  .then(function(r){return r.json();}).then(function(j){
    if(!j||!j.ok)return; _tl.total=j.total||0;
    var el=document.getElementById('tlList'); if(!el)return;
    var out=[],lastD='';
    (j.events||[]).forEach(function(e){
      var d=new Date((e.t||0)*1000), ds=d.toLocaleDateString();
      if(ds!==lastD){ out.push('<div style="color:var(--accent2);font-size:10px;margin:8px 0 4px;border-top:1px dashed rgba(148,163,184,.3);padding-top:6px">🗓 '+esc(ds)+'</div>'); lastD=ds; }
      out.push('<div class="aii"><span>'+(e.e==='app:start'?'🟢':(e.e==='app:close'?'🔴':'•'))+'</span><div>'
        +'<b>'+esc(e.u||e.c||'user')+'</b> — '+esc(dfLbl(e.e))
        +'<div style="color:var(--mut);font-size:9px">'+d.toLocaleTimeString()+'</div></div></div>');
    });
    el.innerHTML=out.length?out.join(''):'<div style="color:var(--mut)">— kuch nahi mila —</div>';
    var pages=Math.max(1,Math.ceil(_tl.total/100));
    var pg=document.getElementById('tlPage'); if(pg)pg.textContent='Page '+(p+1)+' / '+pages;
    var inf=document.getElementById('tlInfo'); if(inf)inf.textContent='— kul '+fmt(_tl.total)+' events'+(q?' · "'+q+'"':'');
    var pv=document.getElementById('tlPrev'); if(pv)pv.disabled=(p<=0);
    var nx=document.getElementById('tlNext'); if(nx)nx.disabled=(p>=pages-1);
    el.scrollTop=0;
  }).catch(function(){});
}
function uaLoad(){ var dt=(document.getElementById('uaDate')||{}).value||'';
  fetch('?admin=1&api=feed'+(dt?'&date='+dt:''),{credentials:'same-origin'})
   .then(function(r){return r.json();}).then(function(j){
    if(!j||!j.ok)return; _ua.rows=j.events||[]; _ua.date=j.date;
    var sel=document.getElementById('uaDate');
    if(sel&&!sel.options.length){ (j.dates&&j.dates.length?j.dates:[j.date]).forEach(function(dd){
      var o=document.createElement('option'); o.value=dd; o.textContent='🗓 '+dd; sel.appendChild(o); }); sel.value=j.date; }
    var us={}; _ua.rows.forEach(function(e){ if(e.u)us[e.u]=1; });
    var su=document.getElementById('uaUser');
    if(su){ var cur=su.value; su.innerHTML='<option value="">👥 Sab users</option>';
      Object.keys(us).sort().forEach(function(k){ var o=document.createElement('option'); o.value=k; o.textContent=k; su.appendChild(o); }); su.value=cur; }
    uaRender(); }).catch(function(){});
}
function _uaRows(){ var fu=(document.getElementById('uaUser')||{}).value||'';
  var ft=(document.getElementById('uaType')||{}).value||'';
  var q=((document.getElementById('uaQ')||{}).value||'').toLowerCase();
  return _ua.rows.filter(function(e){ if(fu&&e.u!==fu)return false;
    if(ft&&uaCat(e.e)!==ft)return false;
    if(q&&(String(e.e)+' '+(e.u||'')).toLowerCase().indexOf(q)<0)return false; return true;});
}
function uaRender(){ if(!document.getElementById('uaKpis'))return;
  var rows=_uaRows().slice().sort(function(a,b){return (a.t||0)-(b.t||0);});
  var fu=(document.getElementById('uaUser')||{}).value||'';
  // ---- KPI summary ----
  var users={},acts={},hours={},opens=0;
  rows.forEach(function(e){ var u=e.u||e.c||'?';
    if(!users[u])users[u]={n:0,acts:{},first:e.t,last:e.t,sess:0,prev:0};
    var U=users[u]; U.n++; U.last=e.t; if(e.t<U.first)U.first=e.t;
    if(e.e==='app:start'||((e.t-U.prev)>1800&&U.prev))U.sess++;
    if(!U.prev)U.sess=Math.max(U.sess,1); U.prev=e.t;
    var lb=dfLbl(e.e); acts[lb]=(acts[lb]||0)+1;
    var h=new Date((e.t||0)*1000).getHours(); hours[h]=(hours[h]||0)+1;
    if(e.e==='app:start')opens++; });
  var ulist=Object.keys(users); var top='—',topn=0;
  Object.keys(acts).forEach(function(k){ if(acts[k]>topn){topn=acts[k];top=k;} });
  var busy='—',busyn=0; ulist.forEach(function(u){ if(users[u].n>busyn){busyn=users[u].n;busy=u;} });
  var kp=document.getElementById('uaKpis');
  if(kp)kp.innerHTML=_kpi('','📜',fmt(rows.length),'Events ('+esc(_ua.date||'')+')')
    +_kpi('g','👥',ulist.length,'Active users')
    +_kpi('p','🏆',esc(busy),'Sabse busy ('+busyn+')')
    +_kpi('y','🔥',esc(top),'Sabse zyada kaam ('+topn+')')
    +_kpi('','🟢',opens,'App kholi (baar)')
    +_kpi('','📊',ulist.length?Math.round(rows.length/ulist.length):0,'Avg events/user');
  // ---- per-user summary table ----
  var uu=document.getElementById('uaUsers');
  if(uu){ var ur=ulist.sort(function(a,b){return users[b].n-users[a].n;}).map(function(u){ var U=users[u];
      var ta='—',tan=0; Object.keys(U.acts||{}).length; // top act per user
      var pacts={}; rows.forEach(function(e){ if((e.u||e.c)===u){ var l=dfLbl(e.e); pacts[l]=(pacts[l]||0)+1; } });
      Object.keys(pacts).forEach(function(k){ if(k.indexOf('App ')===0)return; if(pacts[k]>tan){tan=pacts[k];ta=k;} });
      return '<tr style="cursor:pointer" onclick="document.getElementById(\'uaUser\').value='+JSON.stringify(u).replace(/"/g,'&quot;')+';uaRender()">'
        +'<td><b>'+esc(u)+'</b></td><td>'+U.n+'</td><td style="font-size:10px">'+esc(ta)+'</td>'
        +'<td style="font-size:10px">'+new Date(U.first*1000).toLocaleTimeString()+' → '+new Date(U.last*1000).toLocaleTimeString()+'</td>'
        +'<td>'+Math.max(1,U.sess)+'</td></tr>'; });
    uu.innerHTML=ur.length?_tbl(ur,['User','Events','Sabse zyada','Pehli → aakhri','Sessions']):'<div style="color:var(--mut)">— aaj koi activity nahi —</div>'; }
  // ---- top actions: TABLE — poora naam dikhe, sabse bada number sabse upar ----
  var ta2=Object.keys(acts).map(function(k){return [k,acts[k]];}).sort(function(a,b){return b[1]-a[1];});
  var mx=ta2.length?ta2[0][1]:1;
  var te=document.getElementById('uaTop');
  // (v3.3.4) _tbl nahi — usme 420px scroll-limit hai; poori list bina
  // scroll ke dikhni chahiye
  if(te)te.innerHTML=ta2.length?('<table><thead><tr><th>Kaam (poora naam)</th><th></th></tr></thead><tbody>'
      +ta2.map(function(r){
        return '<tr><td style="white-space:normal;word-break:break-word;font-size:11px">'+esc(r[0])+'</td>'
          +'<td style="width:52px;text-align:right"><b>'+r[1]+'</b></td></tr>'; }).join('')
      +'</tbody></table>')
    :'<div style="color:var(--mut)">—</div>';
  // ---- hour bars ----
  var he=document.getElementById('uaHours');
  if(he){ var hm=1; Object.keys(hours).forEach(function(h){ if(hours[h]>hm)hm=hours[h]; });
    var hh=''; for(var h=0;h<24;h++){ if(!hours[h])continue;
      hh+='<div style="display:flex;gap:8px;align-items:center;margin-bottom:4px;font-size:11px">'
        +'<span style="flex:0 0 56px;white-space:nowrap">'+(h%12||12)+(h<12?' AM':' PM')+'</span>'
        +'<span class="bar" style="flex:0 0 '+Math.max(5,Math.round(hours[h]*60/hm))+'%;height:9px"></span><b>'+hours[h]+'</b></div>'; }
    he.innerHTML=hh||'<div style="color:var(--mut)">—</div>'; }
}
function uaCSV(){ var rows=_uaRows().slice().sort(function(a,b){return (a.t||0)-(b.t||0);});
  var csv='date,time,user,action\n'+rows.map(function(e){ var d=new Date((e.t||0)*1000);
    return d.toLocaleDateString()+','+d.toLocaleTimeString()+',"'+String(e.u||e.c||'').replace(/"/g,'""')+'","'+String(e.e||'').replace(/"/g,'""')+'"';}).join('\n');
  var a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'})); a.download='apnescan-activity-'+(_ua.date||'')+'.csv'; a.click();
  showToast('⬇ Activity CSV download ho gayi','ok'); }
(function(){ var e1=document.getElementById('uaUser'); if(e1)e1.addEventListener('change',uaRender);
  var e4=document.getElementById('uaType'); if(e4)e4.addEventListener('change',uaRender);
  var e2=document.getElementById('uaQ'); if(e2)e2.addEventListener('input',uaRender);
  var e3=document.getElementById('uaDate'); if(e3)e3.addEventListener('change',uaLoad);
  var e5=document.getElementById('tlQ'); if(e5)e5.addEventListener('keydown',function(ev){ if(ev.key==='Enter')tlGo(0); });
  // ---- (v3.3) SPLITTER: beech ki patti kheencho — ek side chhoti, doosri
  // apne aap badi (chaudai yaad bhi rehti hai) ----
  (function(){ var sp=document.getElementById('uaSplit'),gt=document.getElementById('uaGut'),
      lf=document.getElementById('uaLeft'); if(!sp||!gt||!lf)return;
    function narrow(){ return window.innerWidth<900; }
    function applyMode(){ if(narrow()){ sp.style.flexDirection='column'; gt.style.display='none';
        lf.style.flex='0 0 auto'; } else { sp.style.flexDirection='row'; gt.style.display='flex';
        lf.style.flex='0 0 '+(localStorage.getItem('uaSplitPct')||'52')+'%'; } }
    applyMode(); window.addEventListener('resize',applyMode);
    var drag=false;
    function move(x){ var r=sp.getBoundingClientRect();
      var pct=Math.min(75,Math.max(25,(x-r.left)/r.width*100));
      lf.style.flex='0 0 '+pct.toFixed(1)+'%';
      try{localStorage.setItem('uaSplitPct',pct.toFixed(1));}catch(e){} }
    gt.addEventListener('mousedown',function(e){ if(narrow())return; drag=true;
      document.body.style.cursor='col-resize'; e.preventDefault(); });
    window.addEventListener('mousemove',function(e){ if(drag)move(e.clientX); });
    window.addEventListener('mouseup',function(){ drag=false; document.body.style.cursor=''; });
    gt.addEventListener('touchstart',function(e){ if(!narrow())drag=true; },{passive:true});
    window.addEventListener('touchmove',function(e){ if(drag&&e.touches[0])move(e.touches[0].clientX); },{passive:true});
    window.addEventListener('touchend',function(){ drag=false; });
  })();
  if(document.getElementById('uaKpis')){ uaLoad(); tlGo(0); } })();
// ---- SYSTEM HEALTH ----
(function(){ var el=document.getElementById('healthKpis'); if(!el)return;
  var H=D.health||{}, SY=D.sys||{};
  el.innerHTML=_kpi('g','🌐','OK','Server — PHP '+esc(SY.php||D.srv||'?'))
    +_kpi('','⚡',(D.respMs||0)+' ms','API response')
    +(SY.loadPct!==undefined?_kpi(SY.loadPct>90?'r':(SY.loadPct>60?'y':'g'),'💻',SY.loadPct+'%','CPU load ('+SY.load+' / '+SY.cores+' cores)')
      :(SY.load!==undefined?_kpi('','💻',SY.load,'CPU load (server-wide)'):''))
    +(SY.diskFreePct!==undefined?_kpi(SY.diskFreePct<10?'r':(SY.diskFreePct<25?'y':'g'),'💽',SY.diskFreePct+'%','Disk free ('+SY.diskFreeGB+'/'+SY.diskTotalGB+' GB)'):'')
    +_kpi('p','📄',(D.fileKB||0)+' KB','stats.json size')
    +_kpi('y','📁',(H.backupKB||0)+' KB','Backups size')
    +_kpi('','🧠',(SY.memMB||H.memMB||0)+' MB','Memory (peak)')
    +(SY.logKB!==undefined?_kpi(SY.logKB>400?'y':'','📋',SY.logKB+' KB','error.log size'):'')
    +_kpi('g','📚',fmt(H.records||0),'Total records');
  document.getElementById('healthStore').innerHTML = D.health?
    '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;font-size:11.5px">'
    +'<div>💾 Last save<br><b>'+(H.lastSave?new Date(H.lastSave*1000).toLocaleString():'—')+'</b></div>'
    +'<div>📦 Last backup<br><b>'+(H.lastBackup?new Date(H.lastBackup*1000).toLocaleString():'—')+'</b></div>'
    +'<div>♻️ Last recovery<br><b>'+(H.lastRecovery?new Date(H.lastRecovery*1000).toLocaleString():'kabhi zaroorat nahi ✅')+'</b></div>'
    +'<div>🕒 Load / Save<br><b>'+(H.loadMs||0)+' / '+(H.saveMs||0)+' ms</b></div>'
    +'<div>🔒 Locking<br><b>flock (exclusive RMW)</b></div>'
    +'<div>🗄️ Backups<br><b>.bak ×5 + daily ×14</b></div></div>'
    : '<div style="color:var(--mut)">json_storage.php upload karke poori health dekho (System page → Storage module)</div>';
  // (v2) asli backup files + error.log tail
  var SY2=D.sys||{};
  var bkEl=document.getElementById('healthBk');
  if(bkEl) bkEl.innerHTML=(SY2.backups&&SY2.backups.length)?
    _tbl(SY2.backups.map(function(b){return '<tr><td>📦 '+esc(b.n)+'</td><td style="text-align:right">'+b.kb+' KB</td><td style="text-align:right;color:var(--mut)">'+(b.t?ago(b.t):'—')+'</td></tr>';}),['File','Size','When'])
    :'<div style="color:var(--mut)">— koi backup file nahi mili —</div>';
  var lgEl=document.getElementById('healthLog');
  if(lgEl) lgEl.innerHTML=(SY2.logTail&&SY2.logTail.length)?
    '<pre style="max-height:240px;overflow:auto;font-size:10.5px;line-height:1.5;white-space:pre-wrap;margin:0;color:var(--mut)">'+SY2.logTail.map(function(l){return esc(l);}).join('\n')+'</pre>'
    :'<div style="color:var(--mut)">— log khaali hai (koi error nahi) 🎉 —</div>';
})();
// ---- ACTIVITY LOGS ----
(function(){ var el=document.getElementById('logLogins'); if(!el)return;
  var L=D.adminLoginsFull||D.adminLogins||[];
  el.innerHTML = L.length? _tbl(L.map(function(x){ var t=x.t?new Date(x.t*1000).toLocaleString():(x[0]||'—'); var ip=x.ip||x[1]||'?';
    return '<tr><td>🔐 Login</td><td>'+esc(''+ip)+'</td><td style="text-align:right;color:var(--mut)">'+esc(''+t)+'</td></tr>';}),['Event','IP','Time'])
    : '<div style="color:var(--mut)">—</div>';
  var A=D.auditLog||[];
  document.getElementById('logAudit').innerHTML = A.length? _tbl(A.map(function(a){
    return '<tr><td><b>'+esc(a.act)+'</b>'+(a.det?' <span style="color:var(--mut)">'+esc(a.det)+'</span>':'')+'</td><td>'+esc(a.ip||'?')+'</td><td style="text-align:right;color:var(--mut)">'+ago(a.t)+'</td></tr>';}),['Action','IP','When'])
    : '<div style="color:var(--mut)">— abhi koi admin action nahi —</div>';
  var F=D.failLogRecent||[];
  var fe=document.getElementById('logFails');
  if(fe) fe.innerHTML = F.length? _tbl(F.map(function(x){
    return '<tr><td>🚨 Galat password</td><td>'+esc(x.ip||'?')+'</td><td style="text-align:right;color:var(--mut)">'+(x.t?ago(x.t):'—')+'</td></tr>';}),['Event','IP','When'])
    : '<div style="color:var(--mut)">— koi galat koshish nahi 🎉 —</div>';
})();
// ================= COMMAND PALETTE (Ctrl+K) =================
(function(){ var pal=document.createElement('div'); pal.id='cpal';
  pal.innerHTML='<div class="box"><input id="cpq" placeholder="Module ya user ka naam type karo…"><div class="res" id="cpres"></div></div>';
  document.body.appendChild(pal);
  var MODS=[['🏠','Dashboard','overview'],['🟢','Live Monitoring','live'],['👥','Users','users'],['🖨','Devices','hw'],
    ['📄','Scans','scans'],['📈','Analytics','trends'],['📑','Reports','reports'],['❤️','System Health','health'],
    ['📋','Activity Logs','logs'],['🔁','Growth','growth'],['🧰','Tools & Impact','tools'],['💡','Suggestions','ideas'],['🖥','System','system']];
  var q=document.getElementById('cpq'),res=document.getElementById('cpres'),sel=0,items=[];
  function build(){ var v=(q.value||'').toLowerCase(); items=[];
    MODS.forEach(function(m){ if(!v||m[1].toLowerCase().indexOf(v)>-1) items.push({ic:m[0],t:m[1],k:'Module',fn:(function(p){return function(){jumpToPage(p);};})(m[2])}); });
    if(v.length>=2)(D.userList||[]).forEach(function(u){ if(((u.name||'')+' '+(u.cc||'')+' '+(u.version||'')+' '+(u.gip||'')).toLowerCase().indexOf(v)>-1&&items.length<14)
      items.push({ic:u.online?'🟢':'👤',t:u.name+' · '+fmt(u.scans)+' scans',k:'User',fn:(function(id){return function(){ jumpToPage('users'); if(window.openUser)setTimeout(function(){openUser(id);},200);};})(u.id)}); });
    // (v2) GLOBAL DATA SEARCH: scans / crashes / feedback bhi
    if(v.length>=2){
      var seen={};
      (D.recentScans300||[]).forEach(function(x){ if(items.length<16&&x.name&&!seen['s'+x.name]&&(x.name+' '+(x.sm||'')).toLowerCase().indexOf(v)>-1){ seen['s'+x.name]=1;
        items.push({ic:'📄',t:x.name+' · '+x.n+' pages · '+ago(x.t),k:'Scan',fn:function(){jumpToPage('scans');}}); } });
      (D.crashes||[]).forEach(function(c){ if(items.length<18&&((c.err||'')+' '+(c.v||'')).toLowerCase().indexOf(v)>-1)
        items.push({ic:'💥',t:'Crash v'+(c.v||'?')+': '+String(c.err||'').slice(0,45),k:'Crash',fn:function(){jumpTo('system','cg');}}); });
      (D.feedback||[]).forEach(function(f){ if(items.length<20&&((f.msg||'')+' '+(f.name||'')).toLowerCase().indexOf(v)>-1)
        items.push({ic:'💬',t:(f.name||'user')+': '+String(f.msg||'').slice(0,45),k:'Feedback',fn:function(){jumpToPage('system');}}); });
    }
    sel=0; render(); }
  function render(){ res.innerHTML=items.map(function(i,ix){ return '<div class="ri'+(ix===sel?' sel':'')+'" onclick="void(0)" data-ix="'+ix+'"><span>'+i.ic+'</span> '+esc(i.t)+'<span class="k">'+i.k+'</span></div>'; }).join('')||'<div class="ri" style="color:var(--mut)">— kuch nahi mila —</div>';
    [].slice.call(res.children).forEach(function(el){ el.addEventListener('click',function(){ var i=items[parseInt(el.getAttribute('data-ix')||'-1')]; if(i){close();i.fn();} }); }); }
  function open(){ pal.classList.add('open'); q.value=''; build(); setTimeout(function(){q.focus();},30); }
  function close(){ pal.classList.remove('open'); }
  window.cpalOpen=open;
  q.addEventListener('input',build);
  q.addEventListener('keydown',function(e){ if(e.key==='ArrowDown'){sel=Math.min(sel+1,items.length-1);render();e.preventDefault();}
    else if(e.key==='ArrowUp'){sel=Math.max(sel-1,0);render();e.preventDefault();}
    else if(e.key==='Enter'){ if(items[sel]){close();items[sel].fn();} }
    else if(e.key==='Escape')close(); });
  pal.addEventListener('click',function(e){ if(e.target===pal)close(); });
  document.addEventListener('keydown',function(e){ if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='k'){ e.preventDefault(); open(); } });
})();
// ================= TOASTS + FULLSCREEN =================
function showToast(msg,cls){ var c=document.getElementById('toasts');
  if(!c){ c=document.createElement('div'); c.id='toasts'; document.body.appendChild(c); }
  var t=document.createElement('div'); t.className='toast '+(cls||''); t.textContent=msg; c.appendChild(t);
  setTimeout(function(){ t.style.opacity='0'; t.style.transition='opacity .4s'; setTimeout(function(){t.remove();},420); },3800); }
(function(){ var tb=document.querySelector('header .toolbar'); if(!tb)return;
  var fs=document.createElement('button'); fs.className='iconbtn'; fs.title='Full screen'; fs.textContent='⛶';
  fs.onclick=function(){ if(document.fullscreenElement)document.exitFullscreen(); else document.documentElement.requestFullscreen(); };
  tb.insertBefore(fs,tb.querySelector('.prof')); })();
// SU message ko toast me bhi dikhao
(function(){ var b=document.querySelector('.wrap > div[style*="border-radius:10px"]');
  if(b&&/✅|❌|↩/.test(b.textContent)) showToast(b.textContent.slice(0,90), /✅|↩/.test(b.textContent)?'ok':'bad'); })();

// ---- back-to-top FAB ----
(function(){ var f=document.createElement('button'); f.id='fabTop'; f.textContent='↑'; f.title='Top';
  f.onclick=function(){ scrollTo({top:0,behavior:'smooth'}); }; document.body.appendChild(f);
  addEventListener('scroll',function(){ f.classList.toggle('on',(scrollY||0)>500); },{passive:true});
})();

</script>
</div></div>
</body></html><?php exit;
}

// ================= API (App se) =================
$action = isset($_REQUEST['action']) ? $_REQUEST['action'] : 'stats';
$client = isset($_REQUEST['client']) ? $_REQUEST['client'] : '';
$today  = today_str(); $now = time();

// ---- INPUT VALIDATION (spec #9): har aane wali request saaf-suthri ho ----
// Lambaai-caps + control-characters hatao; number-fields sirf int; kharab IP
// wali request seedha reject. (Purane clients par koi asar nahi — sahi data
// waise ka waisa nikal jaata hai.)
function _clean_str($v, $max) {
    $v = (string)$v;
    $v = preg_replace('/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/', '', $v);  // control chars
    return substr(trim($v), 0, $max);
}
foreach (array('u'=>100, 'sm'=>100, 'c'=>8, 'region'=>80, 'city'=>80, 'v'=>16) as $_k => $_max) {
    if (isset($_REQUEST[$_k])) $_REQUEST[$_k] = _clean_str($_REQUEST[$_k], $_max);
}
foreach (array('n','pg','kb','imp','prt','t','sug','len','num','dt','wd') as $_k) {
    if (isset($_REQUEST[$_k]) && $_REQUEST[$_k] !== '' && !preg_match('/^-?\d+$/', (string)$_REQUEST[$_k]))
        $_REQUEST[$_k] = intval($_REQUEST[$_k]);      // integer-only fields
}
$client = _clean_str($client, 64);
$__ip = isset($_SERVER['REMOTE_ADDR']) ? $_SERVER['REMOTE_ADDR'] : '';
if ($__ip !== '' && !filter_var($__ip, FILTER_VALIDATE_IP)) {
    header('Content-Type: application/json'); echo json_encode(array('ok'=>false,'error'=>'bad request')); exit;
}

// Poore read-modify-write par EXCLUSIVE lock (spec #6) — kai clients ek saath
// bhejein to bhi ginti kabhi na uljhe. (Script ke ant par khud khul jaata hai.)
if (!empty($HAS_STORAGE)) storageLockExclusive($DATA_FILE);
$d = load_data($DATA_FILE);

if ($action === 'scan') {
    if ($SECRET !== '' && (!isset($_REQUEST['secret']) || $_REQUEST['secret'] !== $SECRET)) {
        header('Content-Type: application/json'); echo json_encode(array('ok'=>false,'error'=>'bad secret')); exit;
    }
    $n = max(0, min(100, intval(isset($_REQUEST['n'])?$_REQUEST['n']:1)));
    // ---- DUPLICATE PROTECTION (spec #10): wahi user + scanner + pages ka
    // bilkul wahi scan 15 second ke andar dobara aaye (retry/double-fire) to
    // use dobara mat gino. Alag scans par koi asar nahi.
    $__dupKey = md5($client . '|' . (isset($_REQUEST['u'])?$_REQUEST['u']:'') . '|'
                  . (isset($_REQUEST['sm'])?$_REQUEST['sm']:'') . '|' . $n);
    if (!isset($d['dupGuard']) || !is_array($d['dupGuard'])) $d['dupGuard'] = array();
    $__isDup = isset($d['dupGuard'][$__dupKey]) && ($now - intval($d['dupGuard'][$__dupKey])) < 15;
    $d['dupGuard'][$__dupKey] = $now;
    if (count($d['dupGuard']) > 200) $d['dupGuard'] = array_slice($d['dupGuard'], -150, null, true);
    if ($__isDup) {
        header('Content-Type: application/json'); echo json_encode(array('ok'=>true,'dup'=>true)); exit;
    }
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
        // (v3) SCAN DETAIL: dpi / rang / srot / kitne second / profile / size /
        // scanner / version bhi saath me — admin panel ki detailed grid ke liye
        $d['recentScans'][] = array('t'=>$now, 'name'=>substr(isset($_REQUEST['u'])?$_REQUEST['u']:'',0,40),
            'cc'=>substr(isset($_REQUEST['c'])?$_REQUEST['c']:'',0,4), 'n'=>$n,
            'dpi'=>substr(isset($_REQUEST['dpi'])?$_REQUEST['dpi']:'',0,6),
            'col'=>substr(isset($_REQUEST['col'])?$_REQUEST['col']:'',0,10),
            'sz'=>substr(isset($_REQUEST['sz'])?$_REQUEST['sz']:'',0,10),
            'sm'=>substr(isset($_REQUEST['sm'])?$_REQUEST['sm']:'',0,40),
            'dur'=>max(0,min(3600,intval(isset($_REQUEST['dur'])?$_REQUEST['dur']:0))),
            'src'=>substr(isset($_REQUEST['src'])?$_REQUEST['src']:'',0,8),
            'prof'=>substr(isset($_REQUEST['prof'])?$_REQUEST['prof']:'',0,24),
            'rot'=>substr(isset($_REQUEST['rot'])?$_REQUEST['rot']:'',0,8),
            'tess'=>(isset($_REQUEST['tess'])?intval($_REQUEST['tess']):-1),
            'v'=>substr(isset($_REQUEST['v'])?$_REQUEST['v']:'',0,8));
        $d['recentScans'] = array_slice($d['recentScans'], -1000);   // cap (spec #8)
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
            // (v2.2) USER TIMELINE: har user ki aakhri 30 activities (kab kya kiya)
            if (!isset($d['clients'][$client]['ev'])||!is_array($d['clients'][$client]['ev'])) $d['clients'][$client]['ev']=array();
            $d['clients'][$client]['ev'][]=array($now,$feat);
            $d['clients'][$client]['ev']=array_slice($d['clients'][$client]['ev'],-30);
        }
        // (v2.2) GLOBAL ACTIVITY FEED: sab users ke haal ke events (admin ko dikhta hai)
        if (!isset($d['recentEvents'])||!is_array($d['recentEvents'])) $d['recentEvents']=array();
        $d['recentEvents'][]=array('t'=>$now,'client'=>substr((string)$client,0,40),
            'u'=>substr(trim(isset($_REQUEST['u'])?$_REQUEST['u']:''),0,40),'feat'=>$feat);
        $d['recentEvents']=array_slice($d['recentEvents'],-400);
        // (18) feature adoption over time — roz ka feature count
        if(!isset($d['featDaily'])||!is_array($d['featDaily'])) $d['featDaily']=array();
        if(!isset($d['featDaily'][$today])) $d['featDaily'][$today]=array();
        $d['featDaily'][$today][$feat]=intval(isset($d['featDaily'][$today][$feat])?$d['featDaily'][$today][$feat]:0)+1;
        if(count($d['featDaily'])>40){ ksort($d['featDaily']); $d['featDaily']=array_slice($d['featDaily'],-35,null,true); }
        // rename analytics ka privacy-safe meta (sug=autocomplete-pick, len=lambaai, num/dt=pattern)
        if($feat==='rename'){
            if(!isset($d['renameMeta'])||!is_array($d['renameMeta'])) $d['renameMeta']=array('sugY'=>0,'sugN'=>0,'lenSum'=>0,'lenN'=>0,'numY'=>0,'dtY'=>0,'wdSum'=>0);
            if(isset($_REQUEST['sug'])){ if(intval($_REQUEST['sug']))$d['renameMeta']['sugY']++; else $d['renameMeta']['sugN']++; }
            $__ln=intval(isset($_REQUEST['len'])?$_REQUEST['len']:0); if($__ln>0){ $d['renameMeta']['lenSum']+=min($__ln,120); $d['renameMeta']['lenN']++; }
            if(!empty($_REQUEST['num'])) $d['renameMeta']['numY']++;
            if(!empty($_REQUEST['dt']))  $d['renameMeta']['dtY']++;
            $__wd=intval(isset($_REQUEST['wd'])?$_REQUEST['wd']:0); if($__wd>0) $d['renameMeta']['wdSum']+=min($__wd,20);
        }
    }
    // numeric metrics
    if (!isset($d['metrics'])) $d['metrics'] = array();
    $kb = max(0, min(50000000, intval(isset($_REQUEST['kb'])?$_REQUEST['kb']:0)));
    $pg = max(0, min(5000, intval(isset($_REQUEST['pg'])?$_REQUEST['pg']:0)));
    if ($kb) $d['metrics']['kbSaved'] = intval(isset($d['metrics']['kbSaved'])?$d['metrics']['kbSaved']:0) + $kb;
    if ($pg && !empty($_REQUEST['feat'])) { $fk = 'pg_'.substr($_REQUEST['feat'],0,16); $d['metrics'][$fk] = intval(isset($d['metrics'][$fk])?$d['metrics'][$fk]:0) + $pg; }
} else if ($action === 'evbatch') {
    // ---- (v3) HAR-KAAM FEED: app har 60 sec me apne saare chhote-bade
    // kaam (button/menu/nav clicks) EK batch me bhejti hai. Ye ROZ ki
    // alag file me jama hote hain (events-YYYY-MM-DD.jsonl) — poori
    // history hamesha ke liye; recentEvents/user-timeline bhi taaza.
    touch_client($d, $client, $_REQUEST, 0, $now, $today);
    $evs = json_decode(isset($_REQUEST['events'])?$_REQUEST['events']:'', true);
    if (is_array($evs) && $client !== '') {
        $evs = array_slice($evs, 0, 40);
        $edir = __DIR__ . '/apnescan_events';
        if (!is_dir($edir)) { @mkdir($edir, 0755, true);
            @file_put_contents($edir.'/.htaccess', "Require all denied\nDeny from all\n");
            @file_put_contents($edir.'/index.html', ''); }
        $unm = substr(trim(isset($_REQUEST['u'])?$_REQUEST['u']:''),0,40);
        if (!isset($d['recentEvents'])||!is_array($d['recentEvents'])) $d['recentEvents']=array();
        $lines = '';
        foreach ($evs as $e) {
            if (!is_array($e) || count($e) < 2) continue;
            $et = intval($e[0]); if ($et <= 0 || $et > $now + 3600) $et = $now;
            $en = substr(preg_replace('/[\x00-\x1F"\\\\]/','',(string)$e[1]),0,32);
            if ($en === '') continue;
            $lines .= json_encode(array('t'=>$et,'c'=>substr((string)$client,0,40),
                'u'=>$unm,'e'=>$en), JSON_UNESCAPED_UNICODE)."\n";
            $d['recentEvents'][] = array('t'=>$et,'client'=>substr((string)$client,0,40),'u'=>$unm,'feat'=>$en);
            if (isset($d['clients'][$client])) {
                if(!isset($d['clients'][$client]['ev'])||!is_array($d['clients'][$client]['ev']))$d['clients'][$client]['ev']=array();
                $d['clients'][$client]['ev'][] = array($et,$en);
                $d['clients'][$client]['ev'] = array_slice($d['clients'][$client]['ev'],-30);
            }
        }
        $d['recentEvents'] = array_slice($d['recentEvents'],-400);
        if ($lines !== '') @file_put_contents($edir.'/events-'.date('Y-m-d',$now).'.jsonl', $lines, FILE_APPEND|LOCK_EX);
    }
} else if ($action === 'crash') {
    $_cv = substr(isset($_REQUEST['v'])?$_REQUEST['v']:'',0,10);
    $_ce = substr(isset($_REQUEST['err'])?$_REQUEST['err']:'',0,200);
    $d['crashes'][] = array('t'=>$now,'v'=>$_cv,'err'=>$_ce,'client'=>substr($client,0,40));
    $d['crashes'] = array_slice($d['crashes'], -500);    // cap (spec #8)
    tg_send("💥 <b>Crash report</b> (v".$_cv.")\n".htmlspecialchars($_ce));
} else if ($action === 'feedback') {
    $d['feedback'][] = array('t'=>$now,'name'=>substr(isset($_REQUEST['u'])?$_REQUEST['u']:'',0,40),
        'v'=>substr(isset($_REQUEST['v'])?$_REQUEST['v']:'',0,10),'rating'=>max(0,min(5,intval(isset($_REQUEST['rating'])?$_REQUEST['rating']:0))),
        'msg'=>substr(isset($_REQUEST['msg'])?$_REQUEST['msg']:'',0,400),
        'status'=>'open',                           // (22) ticket status: open/resolved
        'client'=>substr($client,0,40));            // reply is user tak pahunchane ke liye
    $d['feedback'] = array_slice($d['feedback'], -500);  // cap (spec #8)
    touch_client($d, $client, $_REQUEST, 0, $now, $today);
}

// import/print kisi bhi action ke saath
$imp = max(0, min(500, intval(isset($_REQUEST['imp'])?$_REQUEST['imp']:0)));
$prt = max(0, min(500, intval(isset($_REQUEST['prt'])?$_REQUEST['prt']:0)));
if ($imp) $d['imports'] = intval($d['imports']) + $imp;
if ($prt) $d['prints']  = intval($d['prints'])  + $prt;

foreach ($d['online'] as $id => $ts) { if ($now - intval($ts) > 86400) unset($d['online'][$id]); }
// (26) server health: har ghante ki request-ginti + roz ki file-size (load/growth trend)
if(!isset($d['reqHours'])||!is_array($d['reqHours'])) $d['reqHours']=array();
$__rhk=date('Y-m-d-H'); $d['reqHours'][$__rhk]=(isset($d['reqHours'][$__rhk])?intval($d['reqHours'][$__rhk]):0)+1;
if(count($d['reqHours'])>80){ ksort($d['reqHours']); $d['reqHours']=array_slice($d['reqHours'],-72,null,true); }
if(!isset($d['sizeDaily'])||!is_array($d['sizeDaily'])) $d['sizeDaily']=array();
if(!isset($d['sizeDaily'][$today]) && is_file($DATA_FILE)) $d['sizeDaily'][$today]=round(filesize($DATA_FILE)/1024,1);
if(count($d['sizeDaily'])>70){ ksort($d['sizeDaily']); $d['sizeDaily']=array_slice($d['sizeDaily'],-60,null,true); }
maybe_backup($DATA_FILE, $d);      // roz ek auto-backup (cron ke bina bhi)
save_data($DATA_FILE, $d);

header('Content-Type: application/json');
echo json_encode(compute_stats($d, $client));
