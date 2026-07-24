<?php
/**
 * =========================================================================
 *  ApneScan Stats — JSON STORAGE LAYER (json_storage.php)
 * =========================================================================
 *  Production-ready, pure-JSON storage. Koi database nahi — sirf stats.json,
 *  lekin bank-grade suraksha ke saath:
 *
 *    - flock() exclusive/shared locking (race-condition proof)
 *    - Atomic writes: tmp file -> verify -> rename (kabhi aadhi file nahi)
 *    - Rotating backups: stats.json.bak ... stats.json.bak4 (max 5)
 *    - Auto-recovery: main file tuti/khaali/gayab -> sabse achhe backup se
 *    - JSON validation har read/write par
 *    - Size caps: recentScans<=1000, feedback<=500, crashes<=500
 *      (daily/monthly/lifetime summaries KABHI delete nahi hote)
 *    - Error + performance log: logs/error.log (>200ms par warning)
 *    - .htaccess auto-generate (stats.json* browser se kabhi na khule)
 *
 *  PUBLIC API (baaki application SIRF yehi use kare):
 *    initializeStorage($file)          loadJson($file [,$exclusive])
 *    saveJson($file,$data[,$force])    safeRead($file)
 *    safeWrite($file,$data)            validateJson($string)
 *    backupJson($file)                 restoreJson($file)
 *    rotateBackups($file)              cleanupJson(&$data)
 *    storageHealth($file)              jsonLog($fn,$msg[,$level])
 *
 *  FUTURE-READY: yehi functions baad me SQLite/MySQL implementation se
 *  badle ja sakte hain — baaki app ko chhuna nahi padega.
 * =========================================================================
 */

// ---- module state (lock handle + timings) --------------------------------
$GLOBALS['_JS'] = array('lock'=>null, 'lockMode'=>'', 'meta'=>array());

/** logs/error.log me ek line: date time | function | file | message */
function jsonLog($fn, $msg, $level = 'ERROR') {
    try {
        $dir = __DIR__ . '/logs';
        if (!is_dir($dir)) @mkdir($dir, 0755, true);
        $f = $dir . '/error.log';
        // self-capping: 500 KB se bada ho jaye to aakhri ~1500 line rakho
        if (is_file($f) && @filesize($f) > 500000) {
            $tail = @file($f); if ($tail) @file_put_contents($f, implode('', array_slice($tail, -1500)), LOCK_EX);
        }
        $line = date('Y-m-d H:i:s') . " | $level | $fn | " . basename(__FILE__) . " | " . str_replace(array("\r","\n"), ' ', (string)$msg) . "\n";
        @file_put_contents($f, $line, FILE_APPEND | LOCK_EX);
    } catch (Throwable $e) { /* logging khud kabhi crash na kare */ }
}

/** >200ms operations ki warning (spec: performance monitoring) */
function _jsPerf($op, $t0) {
    $ms = (microtime(true) - $t0) * 1000.0;
    $GLOBALS['_JS']['meta'][$op . 'Ms'] = round($ms, 1);
    if ($ms > 200) jsonLog($op, sprintf('slow: %.0f ms', $ms), 'WARN');
    return $ms;
}

/** String ko validate-karke array do (invalid -> null). Kabhi exception nahi. */
function validateJson($s) {
    if (!is_string($s)) return null;
    $s = trim($s);
    if ($s === '' || ($s[0] !== '{' && $s[0] !== '[')) return null;
    $d = json_decode($s, true);
    if (json_last_error() !== JSON_ERROR_NONE || !is_array($d)) return null;
    return $d;
}

/** Data ka "kitna bhara" score — recovery me sabse bhara candidate jeete. */
function _jsScore($d) {
    if (!is_array($d)) return -1;
    return intval(isset($d['total']) ? $d['total'] : 0)
         + count(isset($d['clients']) && is_array($d['clients']) ? $d['clients'] : array())
         + count(isset($d['days']) && is_array($d['days']) ? $d['days'] : array());
}

/** Shared-lock ke saath ek file padho aur validate karo (fail -> null). */
function safeRead($file) {
    $t0 = microtime(true);
    if (!is_file($file)) return null;
    $fh = @fopen($file, 'rb');
    if (!$fh) { jsonLog('safeRead', "open fail: $file"); return null; }
    @flock($fh, LOCK_SH);                    // writers ke beech me mat padho
    $s = stream_get_contents($fh);
    @flock($fh, LOCK_UN); @fclose($fh);
    $d = validateJson($s);
    _jsPerf('load', $t0);
    if ($d === null && $s !== false && trim((string)$s) !== '')
        jsonLog('safeRead', 'invalid JSON in ' . basename($file) . ' (' . json_last_error_msg() . ')');
    return $d;
}

/**
 * Exclusive-lock atomic write: tmp likho -> WAPAS PADH KE verify karo ->
 * rename. Kabhi bhi stats.json par seedha likha nahi jaata.
 */
function safeWrite($file, $d) {
    $t0 = microtime(true);
    $json = json_encode($d);
    if ($json === false || strlen($json) < 2) { jsonLog('safeWrite', 'encode fail: ' . json_last_error_msg()); return false; }
    $tmp = $file . '.tmp';
    $fh = @fopen($tmp, 'wb');
    if (!$fh) { jsonLog('safeWrite', "tmp open fail: $tmp"); return false; }
    @flock($fh, LOCK_EX);
    $w = fwrite($fh, $json);
    @fflush($fh);
    @flock($fh, LOCK_UN); @fclose($fh);
    if ($w !== strlen($json)) { @unlink($tmp); jsonLog('safeWrite', 'short write'); return false; }
    // verify: tmp wapas padho — bina verify rename NAHI
    if (validateJson(@file_get_contents($tmp)) === null) { @unlink($tmp); jsonLog('safeWrite', 'verify fail'); return false; }
    @chmod($tmp, 0644);
    if (!@rename($tmp, $file)) {
        @unlink($tmp); jsonLog('safeWrite', "rename fail -> $file"); return false;
    }
    $GLOBALS['_JS']['meta']['lastSave'] = time();
    _jsPerf('save', $t0);
    return true;
}

/** Backup rotation: .bak4 <- .bak3 <- ... <- .bak <- (main copy). Max 5. */
function rotateBackups($file) {
    for ($i = 4; $i >= 1; $i--) {
        $src = $file . '.bak' . ($i == 1 ? '' : ($i - 1));
        $dst = $file . '.bak' . $i;
        if ($i == 1) $src = $file . '.bak';
        if (is_file($src)) @copy($src, $dst);
    }
}

/** Har safal save ke baad: rotation + taaza .bak. */
function backupJson($file) {
    if (!is_file($file)) return false;
    rotateBackups($file);
    $ok = @copy($file, $file . '.bak');
    if ($ok) $GLOBALS['_JS']['meta']['lastBackup'] = time();
    else jsonLog('backupJson', 'copy fail');
    return $ok;
}

/** Sabse achha (sabse bhara) valid backup dhundo: .bak..bak4 + dated backups. */
function restoreJson($file) {
    $best = null; $bs = -1;
    $cands = array($file . '.bak');
    for ($i = 1; $i <= 4; $i++) $cands[] = $file . '.bak' . $i;
    $g = @glob(dirname($file) . '/backups/stats-*.json');
    if ($g) { rsort($g); $cands = array_merge($cands, array_slice($g, 0, 3)); }
    foreach ($cands as $c) {
        $d = validateJson(@is_file($c) ? @file_get_contents($c) : null);
        $s = _jsScore($d);
        if ($s > $bs) { $bs = $s; $best = $d; }
    }
    if ($best !== null) {
        $GLOBALS['_JS']['meta']['lastRecovery'] = time();
        jsonLog('restoreJson', 'recovered from backup (score=' . $bs . ')', 'WARN');
    }
    return $best;
}

/**
 * Size caps (spec #8): sirf rolling lists chhoti hoti hain —
 * daily/monthly/lifetime summaries kabhi nahi katati.
 */
function cleanupJson(&$d) {
    if (isset($d['recentScans']) && is_array($d['recentScans']) && count($d['recentScans']) > 1000)
        $d['recentScans'] = array_slice($d['recentScans'], -1000);
    if (isset($d['feedback']) && is_array($d['feedback']) && count($d['feedback']) > 500)
        $d['feedback'] = array_slice($d['feedback'], -500);
    if (isset($d['crashes']) && is_array($d['crashes']) && count($d['crashes']) > 500)
        $d['crashes'] = array_slice($d['crashes'], -500);
}

/** Ek hi baar: folders + .htaccess (JSON files browser se band). */
function initializeStorage($file) {
    $dir = dirname($file);
    if (!is_dir($dir . '/logs')) @mkdir($dir . '/logs', 0755, true);
    if (!is_dir($dir . '/backups')) @mkdir($dir . '/backups', 0755, true);
    $ht = $dir . '/.htaccess';
    $rules = "# ApneScan stats — data files browser se kabhi na khulein\n"
           . "<Files \"stats.json\">\nRequire all denied\n</Files>\n"
           . "<FilesMatch \"^stats\\.json\\.(bak[0-9]*|tmp|lock)$\">\nRequire all denied\n</FilesMatch>\n"
           . "<FilesMatch \"^stats-.*\\.json$\">\nRequire all denied\n</FilesMatch>\n";
    $cur = is_file($ht) ? @file_get_contents($ht) : '';
    if (strpos((string)$cur, 'stats.json') === false) {
        @file_put_contents($ht, $rules . $cur, LOCK_EX);
    }
    $lg = $dir . '/logs/.htaccess';
    if (!is_file($lg)) @file_put_contents($lg, "Require all denied\n");
    $bg = $dir . '/backups/.htaccess';
    if (!is_file($bg)) @file_put_contents($bg, "Require all denied\n");
}

/**
 * Poore read-modify-write ke liye EXCLUSIVE lock (alag .lock file par, taaki
 * main file ka atomic rename kabhi lock ke aade na aaye). Script khatam hote
 * hi PHP khud release kar deta hai — deadlock ka dar nahi.
 */
function storageLockExclusive($file, $waitSeconds = 5) {
    if (!empty($GLOBALS['_JS']['lock'])) return true;    // pehle se hai
    $fh = @fopen($file . '.lock', 'c');
    if (!$fh) { jsonLog('storageLock', 'lock open fail'); return false; }
    $t0 = microtime(true);
    while (!@flock($fh, LOCK_EX | LOCK_NB)) {
        if (microtime(true) - $t0 > $waitSeconds) {      // itna lamba? -> bina lock aage badho (service > purity)
            jsonLog('storageLock', 'lock timeout after ' . $waitSeconds . 's', 'WARN');
            @fclose($fh); return false;
        }
        usleep(20000);                                    // 20ms
    }
    $GLOBALS['_JS']['lock'] = $fh; $GLOBALS['_JS']['lockMode'] = 'ex';
    return true;
}

function storageUnlock() {
    if (!empty($GLOBALS['_JS']['lock'])) {
        @flock($GLOBALS['_JS']['lock'], LOCK_UN);
        @fclose($GLOBALS['_JS']['lock']);
        $GLOBALS['_JS']['lock'] = null;
    }
}

/**
 * LOAD (recovery ke saath): main -> backups -> default. App kabhi crash nahi
 * karti — tuti file par turant sabse achha backup lag jaata hai.
 * $defaultFn: khaali default structure dene wala callable (stats.php deta hai).
 */
function loadJson($file, $defaultFn = null) {
    $d = safeRead($file);
    if ($d === null) {
        $d = restoreJson($file);
        if ($d !== null) {
            // theek karke wapas likh do taaki agla read seedha main se ho
            safeWrite($file, $d);
        }
    }
    if ($d === null) {
        jsonLog('loadJson', 'no valid data anywhere — fresh default', 'WARN');
        $d = is_callable($defaultFn) ? call_user_func($defaultFn) : array();
    }
    return $d;
}

/**
 * SAVE (poore guards ke saath) — yahi ek raasta hai disk tak.
 *  - monotonic counters (total/imports/prints kabhi peeche nahi)
 *  - non-force: days/clients ka itihaas kabhi wipe nahi hota
 *  - khaali-wipe guard: bhara data 0 se kabhi replace nahi hota
 *  - caps (cleanupJson) -> atomic verify-write -> rotating backup
 */
function saveJson($file, $d, $force = false) {
    $prev = safeRead($file);
    if ($prev === null) $prev = restoreJson($file);
    if (is_array($prev)) {
        foreach (array('total','imports','prints','peakAll') as $k) {
            if (intval(isset($d[$k])?$d[$k]:0) < intval(isset($prev[$k])?$prev[$k]:0)) $d[$k] = intval($prev[$k]);
        }
        if (!$force) {
            if (isset($prev['days']) && is_array($prev['days'])) {
                if (!isset($d['days']) || !is_array($d['days'])) $d['days'] = array();
                foreach ($prev['days'] as $k => $v)
                    if (!isset($d['days'][$k]) || intval($d['days'][$k]) < intval($v)) $d['days'][$k] = intval($v);
            }
            if (isset($prev['clients']) && is_array($prev['clients'])) {
                if (!isset($d['clients']) || !is_array($d['clients'])) $d['clients'] = array();
                foreach ($prev['clients'] as $k => $v) if (!isset($d['clients'][$k])) $d['clients'][$k] = $v;
            }
        }
    }
    if (intval(isset($d['total'])?$d['total']:0) === 0 && _jsScore($prev) > 0) {
        jsonLog('saveJson', 'empty-wipe guard: refused to overwrite real data with zero', 'WARN');
        return false;
    }
    cleanupJson($d);
    $ok = safeWrite($file, $d);
    if ($ok) backupJson($file);
    return $ok;
}

/** Admin dashboard ke liye storage-sehat ka snapshot. */
function storageHealth($file) {
    $bakSize = 0;
    foreach (array('') as $_) { if (is_file($file.'.bak')) $bakSize += filesize($file.'.bak'); }
    for ($i = 1; $i <= 4; $i++) if (is_file($file.'.bak'.$i)) $bakSize += filesize($file.'.bak'.$i);
    $g = @glob(dirname($file) . '/backups/stats-*.json');
    if ($g) foreach ($g as $bf) $bakSize += @filesize($bf);
    $d = safeRead($file);
    $recs = 0;
    if (is_array($d)) foreach (array('clients','recentScans','feedback','crashes','days') as $k)
        $recs += count(isset($d[$k]) && is_array($d[$k]) ? $d[$k] : array());
    $m = $GLOBALS['_JS']['meta'];
    return array(
        'fileKB'   => is_file($file) ? round(filesize($file)/1024, 1) : 0,
        'backupKB' => round($bakSize/1024, 1),
        'memMB'    => round(memory_get_peak_usage(true)/1048576, 1),
        'lastSave'     => isset($m['lastSave']) ? $m['lastSave'] : (is_file($file) ? filemtime($file) : 0),
        'lastBackup'   => isset($m['lastBackup']) ? $m['lastBackup'] : (is_file($file.'.bak') ? filemtime($file.'.bak') : 0),
        'lastRecovery' => isset($m['lastRecovery']) ? $m['lastRecovery'] : 0,
        'records'  => $recs,
        'loadMs'   => isset($m['loadMs']) ? $m['loadMs'] : 0,
        'saveMs'   => isset($m['saveMs']) ? $m['saveMs'] : 0,
    );
}
