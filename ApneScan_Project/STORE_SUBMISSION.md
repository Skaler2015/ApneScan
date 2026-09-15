# ApneScan — Microsoft Store submission (MSIX)

यह guide बताती है कि ApneScan को **Microsoft Store** पर कैसे डालें। एक बार
Store पर आने के बाद: बहुत बड़ा reach **और** SmartScreen warning हमेशा के लिए
खत्म (Store की app भरोसेमंद मानी जाती है)।

MSIX package अपने आप बनता है — GitHub Actions में **"Build MSIX (Microsoft
Store)"** workflow से (देखें `.github/workflows/build-msix.yml`)।

---

## एक बार का सेटअप (~30–45 मिनट)

### 1. Partner Center account बनाओ
- https://partner.microsoft.com/dashboard/registration पर जाओ।
- Individual developer account की **एक बार की fee ~₹1,500 (~$19)** है।
- Registration पूरा करो।

### 2. App का नाम reserve करो
- Partner Center → **Apps and games → New product → MSIX or PWA app**.
- नाम reserve करो: **ApneScan** (उपलब्ध न हो तो "ApneScan Scanner" जैसा)।

### 3. "Product identity" के 3 मान copy करो
- उसी app में → **Product management → Product identity**.
- वहाँ तीन value मिलेंगी:
  - **Package/Identity/Name** → जैसे `12345ApneSoft.ApneScan`
  - **Package/Identity/Publisher** → जैसे `CN=ABCD1234-5678-...`
  - **Package/Properties/PublisherDisplayName** → जैसे `ApneSoft`

### 4. ये 3 मान GitHub में डालो (Variables — Secrets नहीं)
- GitHub repo → **Settings → Secrets and variables → Actions → Variables tab →
  New repository variable**. तीन बनाओ:
  | Variable | Value (Partner Center से) |
  |---|---|
  | `MSIX_IDENTITY_NAME`     | ऊपर की Name |
  | `MSIX_PUBLISHER`         | ऊपर की Publisher (`CN=...`) |
  | `MSIX_PUBLISHER_DISPLAY` | ऊपर की PublisherDisplayName |
- (ये गुप्त नहीं हैं, इसलिए Variables ठीक हैं।)

---

## हर बार MSIX बनाना + अपलोड करना

### 5. Package बनाओ
- GitHub → **Actions → "Build MSIX (Microsoft Store)" → Run workflow**
  (या कोई `v78` जैसा tag push करो — अपने आप बन जाएगा)।
- हो जाने पर उसी run से **Artifacts → `ApneScan-msix`** download करो — अंदर
  `ApneScan.msix` है।

### 6. Store पर upload करो
- Partner Center → अपनी app → **नया submission** शुरू करो।
- **Packages** section में `ApneScan.msix` upload करो।
  - Identity match होने पर यह accept हो जाएगा (इसलिए step 3–4 ज़रूरी हैं)।
  - Store खुद package को sign करता है — आपको कोई certificate नहीं चाहिए।
- **Store listing** भरो — description/keywords/screenshots `MARKETING.md` से
  copy-paste कर सकते हैं।
- **कम से कम 1–8 screenshots** डालो (app के, 1366×768 या बड़े)।
- Age rating का questionnaire भरो (सब "No" — यह एक utility है)।
- **Submit** करो। Microsoft की समीक्षा में आम तौर पर 24–48 घंटे लगते हैं।

---

## अगली versions

हर नई version पर बस:
1. `apnescan.py` में `VERSION` बढ़ाओ (जैसा हमेशा होता है — workflow इसे
   `78.0.0.0` जैसे 4-भाग में बदल देता है)।
2. MSIX workflow फिर से चलाओ → नया `ApneScan.msix` → Partner Center पर नया
   submission।

---

## तकनीकी जानकारी (जिज्ञासा के लिए)

- Package एक **full-trust desktop MSIX** है (`runFullTrust`) — यानी हमारा
  सामान्य PyInstaller EXE ही अंदर है, बस Store के लिए wrap किया हुआ।
- Architecture: **x86** (32-bit) — 64-bit Windows पर भी चलता है।
- Logo assets (`Assets\*.png`) build के समय `apnescan_icon.png` +
  `apnescan_logo.png` से अपने आप बनते हैं (`msix/make_msix_assets.py`)।
- Manifest template: `msix/AppxManifest.xml` (tokens workflow भरता है)।

### आम गलतियाँ
- **Package rejected — identity mismatch:** step 3–4 के मान बिलकुल वैसे ही
  होने चाहिए जैसे Partner Center में हैं।
- **Version पहले से है:** नई submission के लिए `VERSION` बढ़ाना ज़रूरी है।
- **Screenshot नहीं:** Store को कम से कम 1 screenshot चाहिए।
