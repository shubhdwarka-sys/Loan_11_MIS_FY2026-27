# 🚗 TCO Operations System — MIS Automation (v3.0)
### Loan 11 Possible Pvt. Ltd. | FY 2026-27
> **Brand:** The Car Owner (TCO) | Used Car Loan / NBFC | Delhi-NCR

---

## 📋 संक्षिप्त जानकारी (Overview)

यह TCO ऑपरेशन्स के लिए एक पूरी तरह से ऑटोमेटेड Google Sheets MIS (Management Information System) है। यह 6 ब्रांच (Dwarka, Ghaziabad, Karam Pura, Preet Vihar, Pritam Pura, Vikas Puri) के ऑपरेशन्स को हैंडल करता है। [cite: 14]

**Golden Rule:** डेटा सिर्फ एक बार (DM Sheet में) एंटर होगा और वहां से पूरी MIS में अपने आप सिंक (Auto-propagate) हो जाएगा। मैन्युअल डुप्लीकेट एंट्री की कोई ज़रूरत नहीं!

---

## 🌟 नया क्या है v3.0 में (Latest Updates)

- **Smart RTO Code Generation:** अब `REG NO` में अगर यूज़र स्पेस या डैश लगा दे (जैसे "DL 01" या "DL-01"), तो भी स्क्रिप्ट उसे क्लीन करके सही RTO कोड (e.g. `DL01`) जनरेट करती है।
- **Accurate Pending Days:** RTO Pending Days अब Account Sheet की `DISB RECV DATE` से कैलकुलेट होते हैं (लोन डिस्बर्समेंट के दिन से), न कि सिर्फ फाइल लॉग होने के दिन से। [cite: 31]
- **Smart Auto-Sort:** Account और RTO शीट्स अब नई एंट्री आने पर अपने आप `DM NO` की सीरीज़ (A to Z) में सॉर्ट हो जाती हैं।
- **Strict Data Sync:** डीलर का नाम और डिटेल्स अब सटीक `DM NO` मैच करके ही Account और RTO शीट में सिंक होते हैं (Row मिसमैच का एरर फिक्स)।

---

## 📁 फाइल स्ट्रक्चर (File Structure)

TCO-Operations-System/
│
├── TCO_Operations.gs       ← Main Brain: Data Sync, Sorting, Pending Days calc.
├── TCO_UserControl.gs      ← Admin Menu: Backup, Restore, Sheet Protection
├── TCO_SheetSetup.gs       ← UI Builder: Colors, Headers, Column Widths
├── TCO_DataReset.gs        ← Financial Year / Data Reset utilities
└── README.md               ← Documentation [cite: 14]

---

## 🗂️ शीट का स्ट्रक्चर (Sheet Architecture)

### 🟢 विज़िबल शीट्स (Data Entry)
| Sheet Name | Columns | Purpose |
|------------|---------|---------|
| `DM_DISBURSEMENT MEMO` | 57 | Main case entry (Customer, Loan, Payout details) [cite: 5] |
| `ACCOUNT_PAYMENT_TRACKER` | 45 | Payment tracking (Disbursements, Payouts) [cite: 10] |
| `RTO_TRACKER` | 30 | RC Transfer tracking (Vendors, Pending Days) [cite: 12] |

### 🔴 हिडन शीट्स (Support & System Data)
| Sheet Name | Purpose |
|------------|---------|
| `MASTER_DATA` | Consolidated data from all 3 sheets (For Power BI / n8n) [cite: 32] |
| `DEALER_MASTER` | Authorized Dealer database |
| `TCO_EMPLOYEE_MASTER` | Employee HR records & Branch mapping |
| `_BACKUP_DM / ACC / RTO` | System generated auto-backup sheets |

---

## ⚡ ऑटोमेशन फीचर्स (Smart Automations)

### Auto-Generation (अपने आप बनने वाला डेटा)
- **MONTH:** `DM DATE` से महीना और साल (MM-YYYY) अपने आप बनता है। [cite: 29]
- **RTO CODE:** `REG NO` से RTO कोड निकलता है (e.g., `DL13CR7788` → `DL13`)। [cite: 31]
- **PENDING DAYS:** डिस्बर्समेंट डेट से आज तक के दिन कैलकुलेट करता है। `RC TRANSFER STATUS` 'Done' होने पर रुक जाता है। [cite: 31]
- **RTO PROFIT:** Account शीट में `RTO CHARGES - RTO PAID AMOUNT` कैलकुलेट होता है। [cite: 30]

### Auto-Pull (Cross-Sheet Sync)
- **New Row:** DM शीट में `DATE` और `BUYER NAME` डालते ही Account और RTO शीट में नई Row बन जाती है।
- **Employee Sync:** `EMP ID` डालने पर Executive Name, Branch, और Team Leader अपने आप आ जाते हैं। [cite: 29]
- **Dealer Sync:** `CONTACT NO` डालने पर Dealer Name, Auth Person, और Location आ जाते हैं। [cite: 29]
- **Reverse Sync:** Account शीट की `DISB RECV DATE` और `DEALERSHIP PAYMENT STATUS` वापस DM और RTO शीट में लाइव अपडेट होते हैं। [cite: 30]

---

## 🎨 कलर और स्टाइल प्रणाली (Cell Styles)

यूज़र्स को गाइड करने के लिए कॉलम्स का कलर-कोड फिक्स किया गया है: [cite: 3, 4]

| Type | Indicator | Background Color | Style | Edit Access |
|------|-----------|------------------|-------|-------------|
| **MANUAL** | ✏️ MANUAL | White (`#FFFFFF`) | Normal, Black | Users & Admin |
| **DROPDOWN** | ▾ DROPDOWN | Light Green | Normal, Black | Users & Admin |
| **AUTO PULL**| ↺ AUTO PULL | Light Blue (`#EBF2FF`) | Italic, Blue | **Locked** (Script) |
| **AUTO GEN** | ⚙ AUTO GEN | Light Yellow (`#FFFDE7`) | Italic, Brown | **Locked** (Script) |

*(लाल (Red) या पीला (Yellow) कलर अगर सेल में आता है, तो वह गलत डेटा का वार्निंग साइन है।*

---

## 🔒 सिक्योरिटी और कंट्रोल्स (Protection System)

Google Sheet में **"TCO Admin"** नाम का एक कस्टम मेनू (Custom Menu) है, जो सिर्फ ऑथराइज़्ड एडमिन को दिखता है:

| Menu Item | Action |
|-----------|--------|
| **📦 Manual Backup** | एक क्लिक में DM, ACC, RTO शीट्स का बैकअप ले लेता है। |
| **🔁 Restore** | पुरानी डिलीट हुई या खराब हुई फाइल्स को बैकअप से वापस लाता है। |
| **🔒 Apply Protections** | सभी Auto/Gen कॉलम्स और Hidden शीट्स को लॉक कर देता है। |
| **🎨 Borders & Formatting** | अगर कोई बॉर्डर बिगड़ जाए, तो उसे वापस सेट कर देता है। |
| **🗑️ Data Reset** | नया FY शुरू होने पर पुराना डेटा डिलीट करने के काम आता है। |

---

## 🚀 सेटअप कैसे करें (First Time Setup)

1. अपनी Google Sheet खोलें -> **Extensions** -> **Apps Script** में जाएं।
2. `.gs` फाइल्स का सारा कोड वहां अलग-अलग फाइल्स बनाकर पेस्ट करें।
3. **Triggers सेट करें:**
   - `onEditTrigger` -> On Edit
   - `autoUpdateMaster` -> Time-driven (Every 5 minutes)
   - `dailyPendingDaysUpdate` -> Time-driven (Daily at 8:00 AM)
4. शीट में वापस आएं और **TCO Admin** मेनू से "Apply Borders" और "Apply Protections" रन कर दें।
5. `TCO_EMPLOYEE_MASTER` और `DEALER_MASTER` शीट में अपना बेसिक डेटा डाल दें। बस, सिस्टम रेडी है!


---
> *Built with Google Apps Script | Optimized for Data Analysts & MIS Managers | Ready for n8n API & Power BI Dashboards*
