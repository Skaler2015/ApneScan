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


// ===================== SELF-HOSTED QR GENERATOR =====================
/**
 * Tiny self-contained QR generator (byte mode, ECC level M, versions 1-6).
 * Enough for short URLs (~106 bytes). Returns a boolean matrix (true=dark).
 * No GD / no external service — pure PHP.
 */

// ---- ECC (level M) block table for versions 1..6 ----
//  [ ecc_codewords_per_block, num_blocks, data_codewords_per_block ]
$QR_M = array(
    1 => array(10, 1, 16),
    2 => array(16, 1, 28),
    3 => array(26, 1, 44),
    4 => array(18, 2, 32),
    5 => array(24, 2, 43),
    6 => array(16, 4, 27),
);
// remainder bits per version (1..6)
$QR_REM = array(1=>0, 2=>7, 3=>7, 4=>7, 5=>7, 6=>7);
// alignment-pattern centre (single, at (X,X)) for versions 2..6; v1 = none
$QR_ALIGN = array(2=>18, 3=>22, 4=>26, 5=>30, 6=>34);

// ---- GF(256) tables (primitive poly 0x11d) ----
function qr_gf() {
    static $exp = null, $log = null;
    if ($exp !== null) return array($exp, $log);
    $exp = array_fill(0, 512, 0); $log = array_fill(0, 256, 0);
    $x = 1;
    for ($i = 0; $i < 255; $i++) {
        $exp[$i] = $x; $log[$x] = $i;
        $x <<= 1;
        if ($x & 0x100) $x ^= 0x11d;
    }
    for ($i = 255; $i < 512; $i++) $exp[$i] = $exp[$i - 255];
    return array($exp, $log);
}
function qr_rs_gen($n) {
    list($exp, $log) = qr_gf();
    $g = array(1);
    for ($i = 0; $i < $n; $i++) {
        $ng = array_fill(0, count($g) + 1, 0);
        foreach ($g as $j => $c) {
            $ng[$j] ^= $c;
            $ng[$j + 1] ^= ($c ? $exp[($log[$c] + $i) % 255] : 0);
        }
        $g = $ng;
    }
    return $g;
}
function qr_rs_ecc($data, $n) {
    list($exp, $log) = qr_gf();
    $g = qr_rs_gen($n);
    $res = array_fill(0, $n, 0);
    foreach ($data as $d) {
        $factor = $d ^ $res[0];
        array_shift($res); $res[] = 0;
        if ($factor) {
            for ($i = 0; $i < $n; $i++) {
                $res[$i] ^= $exp[($log[$g[$i + 1]] + $log[$factor]) % 255];
            }
        }
    }
    return $res;
}

function qr_bits_to_codewords($bytes, $version) {
    global $QR_M;
    list($eccpb, $nb, $dpb) = $QR_M[$version];
    $total_data = $nb * $dpb;
    // bit stream: mode(0100) + count(8 bits, v1-9) + data
    $bits = array();
    $push = function($val, $len) use (&$bits) {
        for ($i = $len - 1; $i >= 0; $i--) $bits[] = ($val >> $i) & 1;
    };
    $push(0x4, 4);                       // byte mode
    $push(count($bytes), 8);             // char count (versions 1-9 byte = 8 bits)
    foreach ($bytes as $b) $push($b, 8);
    // terminator (up to 4 zero bits)
    $cap = $total_data * 8;
    for ($i = 0; $i < 4 && count($bits) < $cap; $i++) $bits[] = 0;
    // pad to byte boundary
    while (count($bits) % 8 !== 0) $bits[] = 0;
    // to codewords
    $cw = array();
    for ($i = 0; $i < count($bits); $i += 8) {
        $v = 0; for ($j = 0; $j < 8; $j++) $v = ($v << 1) | $bits[$i + $j];
        $cw[] = $v;
    }
    // pad codewords 0xEC 0x11 ...
    $padvals = array(0xEC, 0x11); $pi = 0;
    while (count($cw) < $total_data) { $cw[] = $padvals[$pi % 2]; $pi++; }
    return $cw;
}

function qr_assemble($version, $bytes) {
    global $QR_M;
    list($eccpb, $nb, $dpb) = $QR_M[$version];
    $cw = qr_bits_to_codewords($bytes, $version);
    // split into blocks (uniform)
    $blocks = array(); $eccblocks = array();
    for ($b = 0; $b < $nb; $b++) {
        $blk = array_slice($cw, $b * $dpb, $dpb);
        $blocks[] = $blk;
        $eccblocks[] = qr_rs_ecc($blk, $eccpb);
    }
    // interleave data then ecc
    $out = array();
    for ($i = 0; $i < $dpb; $i++) foreach ($blocks as $blk) $out[] = $blk[$i];
    for ($i = 0; $i < $eccpb; $i++) foreach ($eccblocks as $eb) $out[] = $eb[$i];
    return $out;
}

function qr_choose_version($len) {
    global $QR_M;
    for ($v = 1; $v <= 6; $v++) {
        list($eccpb, $nb, $dpb) = $QR_M[$v];
        $total_data = $nb * $dpb;
        $avail = $total_data - 1 - 1;   // minus mode+count byte overhead (~2 bytes)
        if ($len <= $avail) return $v;
    }
    return 6;
}

// build the module matrix (true=dark). Returns array of rows.
function qr_matrix($text, &$chosen_mask = null, &$chosen_version = null) {
    global $QR_REM, $QR_ALIGN;
    $bytes = array_values(unpack('C*', $text));
    $version = qr_choose_version(count($bytes));
    $chosen_version = $version;
    $size = 17 + 4 * $version;
    $data = qr_assemble($version, $bytes);

    // module + reserved maps
    $m = array(); $fn = array();   // $fn[r][c]=true => function/reserved (not data)
    for ($r = 0; $r < $size; $r++) { $m[$r] = array_fill(0, $size, 0); $fn[$r] = array_fill(0, $size, false); }

    $setFinder = function($r0, $c0) use (&$m, &$fn) {
        for ($r = -1; $r <= 7; $r++) for ($c = -1; $c <= 7; $c++) {
            $rr = $r0 + $r; $cc = $c0 + $c;
            if ($rr < 0 || $cc < 0 || $rr >= count($m) || $cc >= count($m)) continue;
            $inRing = ($r >= 0 && $r <= 6 && ($c == 0 || $c == 6)) ||
                      ($c >= 0 && $c <= 6 && ($r == 0 || $r == 6));
            $inCore = ($r >= 2 && $r <= 4 && $c >= 2 && $c <= 4);
            $m[$rr][$cc] = ($inRing || $inCore) ? 1 : 0;
            $fn[$rr][$cc] = true;
        }
    };
    $setFinder(0, 0); $setFinder(0, $size - 7); $setFinder($size - 7, 0);

    // timing patterns
    for ($i = 8; $i < $size - 8; $i++) {
        $v = ($i % 2 == 0) ? 1 : 0;
        if (!$fn[6][$i]) { $m[6][$i] = $v; $fn[6][$i] = true; }
        if (!$fn[$i][6]) { $m[$i][6] = $v; $fn[$i][6] = true; }
    }
    // dark module
    $m[4 * $version + 9][8] = 1; $fn[4 * $version + 9][8] = true;

    // alignment pattern (single, versions 2..6)
    if (isset($QR_ALIGN[$version])) {
        $ac = $QR_ALIGN[$version];
        for ($r = -2; $r <= 2; $r++) for ($c = -2; $c <= 2; $c++) {
            $rr = $ac + $r; $cc = $ac + $c;
            $ring = (abs($r) == 2 || abs($c) == 2);
            $centre = ($r == 0 && $c == 0);
            $m[$rr][$cc] = ($ring || $centre) ? 1 : 0;
            $fn[$rr][$cc] = true;
        }
    }

    // reserve format-info areas
    for ($i = 0; $i < 9; $i++) {
        if (!$fn[8][$i]) $fn[8][$i] = true;         // row 8
        if (!$fn[$i][8]) $fn[$i][8] = true;         // col 8
    }
    for ($i = 0; $i < 8; $i++) {
        $fn[8][$size - 1 - $i] = true;
        $fn[$size - 1 - $i][8] = true;
    }

    // ---- place data bits (zigzag) ----
    $bitstream = array();
    foreach ($data as $cw) for ($i = 7; $i >= 0; $i--) $bitstream[] = ($cw >> $i) & 1;
    foreach (range(1, $QR_REM[$version]) as $_) $bitstream[] = 0;   // remainder bits
    $bi = 0;
    $col = $size - 1;
    $upward = true;
    while ($col > 0) {
        if ($col == 6) $col--;      // skip vertical timing
        for ($t = 0; $t < $size; $t++) {
            $row = $upward ? ($size - 1 - $t) : $t;
            for ($k = 0; $k < 2; $k++) {
                $cc = $col - $k;
                if ($fn[$row][$cc]) continue;
                $bit = ($bi < count($bitstream)) ? $bitstream[$bi] : 0;
                $bi++;
                $m[$row][$cc] = $bit;
            }
        }
        $col -= 2; $upward = !$upward;
    }

    // ---- masking: choose best of 8 ----
    $maskFn = function($mask, $r, $c) {
        switch ($mask) {
            case 0: return ($r + $c) % 2 == 0;
            case 1: return $r % 2 == 0;
            case 2: return $c % 3 == 0;
            case 3: return ($r + $c) % 3 == 0;
            case 4: return (intdiv($r, 2) + intdiv($c, 3)) % 2 == 0;
            case 5: return (($r * $c) % 2) + (($r * $c) % 3) == 0;
            case 6: return ((($r * $c) % 2) + (($r * $c) % 3)) % 2 == 0;
            case 7: return ((($r + $c) % 2) + (($r * $c) % 3)) % 2 == 0;
        }
        return false;
    };

    $bestMask = 0; $bestPen = PHP_INT_MAX; $bestMat = null;
    for ($mask = 0; $mask < 8; $mask++) {
        $mm = $m;
        for ($r = 0; $r < $size; $r++) for ($c = 0; $c < $size; $c++) {
            if (!$fn[$r][$c] && $maskFn($mask, $r, $c)) $mm[$r][$c] ^= 1;
        }
        qr_place_format($mm, $fn, $version, $mask, $size);
        $pen = qr_penalty($mm, $size);
        if ($pen < $bestPen) { $bestPen = $pen; $bestMask = $mask; $bestMat = $mm; }
    }
    $chosen_mask = $bestMask;
    return $bestMat;
}

function qr_place_format(&$m, $fn, $version, $mask, $size) {
    // format bits: level M = 0b00, mask 3 bits; BCH(15,5) gen 0x537, xor 0x5412
    $data = (0b00 << 3) | $mask;              // 5-bit
    $code = $data << 10;
    for ($i = 14; $i >= 10; $i--) {
        if (($code >> $i) & 1) $code ^= (0x537 << ($i - 10));
    }
    $bits = (($data << 10) | ($code & 0x3FF)) ^ 0x5412;   // 15 bits
    $count = $size;
    // EXACT placement as per QR spec (matches the reference encoder).
    // vertical (col 8)
    for ($i = 0; $i < 15; $i++) {
        $mod = ($bits >> $i) & 1;
        if ($i < 6)       $m[$i][8] = $mod;
        elseif ($i < 8)   $m[$i + 1][8] = $mod;
        else              $m[$count - 15 + $i][8] = $mod;
    }
    // horizontal (row 8)
    for ($i = 0; $i < 15; $i++) {
        $mod = ($bits >> $i) & 1;
        if ($i < 8)       $m[8][$count - $i - 1] = $mod;
        elseif ($i < 9)   $m[8][15 - $i - 1 + 1] = $mod;   // = col 7
        else              $m[8][15 - $i - 1] = $mod;
    }
}

function qr_penalty($m, $size) {
    $pen = 0;
    // Rule 1: runs of 5+ same colour in rows/cols
    for ($r = 0; $r < $size; $r++) {
        $run = 1;
        for ($c = 1; $c < $size; $c++) {
            if ($m[$r][$c] == $m[$r][$c-1]) { $run++; }
            else { if ($run >= 5) $pen += 3 + ($run - 5); $run = 1; }
        }
        if ($run >= 5) $pen += 3 + ($run - 5);
    }
    for ($c = 0; $c < $size; $c++) {
        $run = 1;
        for ($r = 1; $r < $size; $r++) {
            if ($m[$r][$c] == $m[$r-1][$c]) { $run++; }
            else { if ($run >= 5) $pen += 3 + ($run - 5); $run = 1; }
        }
        if ($run >= 5) $pen += 3 + ($run - 5);
    }
    // Rule 2: 2x2 blocks
    for ($r = 0; $r < $size-1; $r++) for ($c = 0; $c < $size-1; $c++) {
        $v = $m[$r][$c];
        if ($m[$r][$c+1]==$v && $m[$r+1][$c]==$v && $m[$r+1][$c+1]==$v) $pen += 3;
    }
    // Rule 3: finder-like 1011101 with 4 light on a side
    $pat1 = array(1,0,1,1,1,0,1,0,0,0,0);
    $pat2 = array(0,0,0,0,1,0,1,1,1,0,1);
    for ($r = 0; $r < $size; $r++) for ($c = 0; $c <= $size-11; $c++) {
        $ok1 = true; $ok2 = true;
        for ($k = 0; $k < 11; $k++) { if ($m[$r][$c+$k] != $pat1[$k]) $ok1 = false; if ($m[$r][$c+$k] != $pat2[$k]) $ok2 = false; }
        if ($ok1 || $ok2) $pen += 40;
    }
    for ($c = 0; $c < $size; $c++) for ($r = 0; $r <= $size-11; $r++) {
        $ok1 = true; $ok2 = true;
        for ($k = 0; $k < 11; $k++) { if ($m[$r+$k][$c] != $pat1[$k]) $ok1 = false; if ($m[$r+$k][$c] != $pat2[$k]) $ok2 = false; }
        if ($ok1 || $ok2) $pen += 40;
    }
    // Rule 4: dark proportion
    $dark = 0; $tot = $size * $size;
    for ($r = 0; $r < $size; $r++) for ($c = 0; $c < $size; $c++) $dark += $m[$r][$c];
    $ratio = $dark * 100 / $tot;
    $prev = floor(abs($ratio - 50) / 5) * 5;
    $pen += (int)($prev * 2);
    return $pen;
}

// Render QR as a self-contained SVG (white bg + 4-module quiet zone).
function qr_svg($text, $px = 232) {
    $mk = null; $ver = null;
    $mat = qr_matrix($text, $mk, $ver);
    $n = count($mat); $q = 4; $dim = $n + 2 * $q;
    $d = '';
    for ($r = 0; $r < $n; $r++) for ($c = 0; $c < $n; $c++) {
        if ($mat[$r][$c]) { $x = $c + $q; $y = $r + $q; $d .= "M{$x} {$y}h1v1h-1z"; }
    }
    return '<svg xmlns="http://www.w3.org/2000/svg" width="' . $px . '" height="' . $px . '" '
        . 'viewBox="0 0 ' . $dim . ' ' . $dim . '" shape-rendering="crispEdges" '
        . 'style="border-radius:12px">'
        . '<rect width="' . $dim . '" height="' . $dim . '" fill="#fff"/>'
        . '<path d="' . $d . '" fill="#000"/></svg>';
}
// ===================================================================

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
    <?php echo qr_svg($base . '?r=' . $room, 230); ?>
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
