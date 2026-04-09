# 🚗 TCO Operations System — MIS Automation
### Loan 11 Possible Pvt. Ltd. | FY 2026-27
> Brand: **The Car Owner (TCO)** | Used Car Loan / NBFC | Delhi-NCR

---

## 📋 Overview

A fully automated Google Sheets MIS (Management Information System) for used car loan operations across 6 branches: **Dwarka, Ghaziabad, Karam Pura, Preet Vihar, Pritam Pura, Vikas Puri**.

**Golden Rule:** Data entered once → auto-propagates everywhere. No duplicate manual entry.

---

## 📁 File Structure

```
TCO-Operations-System/
│
├── TCO_Operations.gs       ← Main automation script (Admin account)
├── TCO_UserControl.gs      ← Sheet protection & access control
├── TCO_SheetSetup.gs       ← One-time sheet setup & formatting
├── TCO_DataReset.gs        ← Data reset utilities
└── README.md               ← This file
```

---

## 🗂️ Sheet Architecture

### Visible Sheets (3) — Data Entry
| Sheet | Columns | Purpose |
|-------|---------|---------|
| `DM_DISBURSEMENT MEMO` | 57 | Main case entry — loans, customers, payouts |
| `ACCOUNT_PAYMENT_TRACKER` | 45 | Payment tracking — disbursements, payouts |
| `RTO_TRACKER` | 30 | RC transfer tracking — vendors, pending days |

### Hidden Sheets (7) — Support Data
| Sheet | Purpose |
|-------|---------|
| `MASTER_DATA` | Consolidated data from all 3 sheets |
| `DEALER_MASTER` | Authorized dealer database |
| `TCO_EMPLOYEE_MASTER` | Employee records |
| `AUTH_AUDIT_LOG` | Login history |
| `_BACKUP_DM` | DM sheet backup |
| `_BACKUP_ACCOUNT` | Account sheet backup |
| `_BACKUP_RTO` | RTO sheet backup |

---

## ⚡ Automation Features

### Auto-Generation
- **DM_NO** → Format: `TCOAPR26001` (auto on DM DATE + BUYER NAME entry)
- **MONTH** → MM-YYYY from DM DATE
- **NO OF DAYS** → Formula: DM DATE to today (stops on RC TRANSFER DATE)
- **RTO CODE** → From Registration No (e.g. `DL13CR7788` → `DL13`)
- **PENDING DAYS** → Auto-calculates, stops when RC date entered

### Auto-Pull (Cross-Sheet Sync)
- DM entry → ACC + RTO rows created automatically
- EMP ID → Executive Name, Branch, Team Leader auto-fill
- Contact No → Dealer Name, Auth Person, Location auto-fill
- ACC DEALER PAY STATUS → syncs to DM + RTO
- ACC DISB RECV DATE → syncs back to DM as DM_RECV_DATE
- RTO VENDOR → syncs to ACC
- RC STATUS → syncs to ACC

### Validations
| Type | Rule |
|------|------|
| Dropdown | Strict block (Bank, Product, Status) |
| Date | Strict block (`d-mmm-yyyy` format) |
| Tenure | 2-digit max |
| MFG Year | 4-digit (1900-9999) |
| Currency | Warning only (red cell) |
| Phone | Warning only (red cell) |
| Email | Lowercase, warning if no `@` |

---

## 🎨 Cell Style System

| Type | Color | Style |
|------|-------|-------|
| AUTO GEN | `#166534` Dark Green | Italic, size 10 |
| AUTO PULL | `#1A4D8F` Deep Blue | Italic, size 10 |
| MANUAL | `#000000` Black | Normal, size 10 |
| DM_NO | `#5D4037` Dark Brown | Italic, size 12 |
| REG NO | `#c2410c` Dark Orange | Normal, size 12 |
| NOT FOUND | `#ef4444` Red | Italic, size 10 |
| PENDING DAYS | `#ef4444` Red | Normal, size 12 |

---

## 🔒 Protection System

```
8 Hidden Sheets  → Locked (Admin only)
3 Visible Sheets → Open for editors
Auto Cols        → Range-locked (Admin only)
Manual Cols      → Editable by assigned users
```

### TCO_UserControl.gs Functions
| Function | Action |
|----------|--------|
| `TCO_lockHiddenSheets()` | Hide + lock 7 support sheets |
| `TCO_lockAutoCols()` | Lock auto-generated columns |
| `TCO_applyAllProtections()` | Run both above together |
| `TCO_removeProtections()` | Emergency — remove all protections |

### Auto-Locked Columns
| Sheet | Auto-Locked Cols |
|-------|-----------------|
| DM | 1, 3, 4, 5, 39, 41, 42, 43, 44, 45, 47, 48 |
| Account | 1-14, 39, 40, 43, 44 |
| RTO | 1-8, 10, 11, 12, 15, 17, 18, 19, 26 |

---

## 🚀 Setup Instructions

### First Time Setup

**Step 1 — Admin account (your-admin@gmail.com):**
```
Google Sheet → Extensions → Apps Script
→ Add files: TCO_Operations.gs, TCO_UserControl.gs,
             TCO_SheetSetup.gs, TCO_DataReset.gs
→ Add Triggers:
   onEdit  → From Spreadsheet → On edit
   onOpen  → From Spreadsheet → On open
```

**Step 2 — Run Sheet Setup:**
```
Google Sheet → TCO Admin menu → Run individually:
1. createAllSheets_
2. setDmHeaders_
3. setAccHeaders_
4. setRtoHeaders_
5. setValidations_
6. applyAllFormatting_
```

**Step 3 — Apply Protections:**
```
TCO Admin menu → 🔒 Apply Protections
```

**Step 4 — Fill Master Data:**
```
TCO_EMPLOYEE_MASTER → Add employee records
DEALER_MASTER → Add dealer records
```

---

## 📋 TCO Admin Menu

| Menu Item | Function |
|-----------|----------|
| 📦 Manual Backup | Backup DM, ACC, RTO sheets |
| 🔒 Apply Protections | Lock auto cols + hidden sheets |
| 🗂️ Borders & Formatting | Reapply cell borders |
| 📋 Update Master Data | Sync master data sheet |
| 🔁 Restore from Backup | Restore DM from backup |
| 🗑️ Data Reset | Reset individual or all sheets |

---

## 👥 User Access

| Role | DM | ACC | RTO |
|------|----|-----|-----|
| Admin/Owner | ✅ Full | ✅ Full | ✅ Full |
| Data Entry | ✏️ Manual cols | ✏️ Manual cols | ✏️ Manual cols |
| View Only | 👁️ View | 👁️ View | 👁️ View |

> Users added as **Viewer** in Google Sheet cannot see Extensions/Apps Script.
> Manual columns unlocked per user via Google Sheets → Data → Protect Sheets & Ranges.

---

## 🔧 Key Configuration

```javascript
// In TCO_Operations.gs
var TCO = {
  SHEET_ID    : 'YOUR_SHEET_ID_HERE',
  DATA_ROW    : 4,           // Data starts row 4
  ADMIN_EMAIL : 'your-admin@gmail.com'
}
```

---

## 📌 Important Notes

- **Row 1** = Section labels (colored headers)
- **Row 2** = Column headings
- **Row 3** = Type indicators (AUTO GEN / AUTO PULL / MANUAL)
- **Row 4+** = Data entry rows


---

## 🏢 Branch Codes
`DWARKA` | `GHAZIABAD` | `KARAM PURA` | `PREET VIHAR` | `PRITAM PURA` | `VIKAS PURI`

---

## 📞 Contact
**Admin:** your-admin@gmail.com
**System:** Loan 11 Possible Pvt. Ltd. — The Car Owner (TCO)
**FY:** 2026-27

---
*Built with Google Apps Script | Zero external dependencies | Free deployment*
