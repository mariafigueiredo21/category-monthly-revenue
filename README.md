# 💰 Monthly Revenue Automation

This project automates the entire process of **updating the monthly revenue records** in Excel, ensuring data consistency, formatting alignment, and automatic category mapping — all within a few minutes.

Originally developed for internal operations, all files, names, and directories have been **fully anonymized** for public release.

---

## ⚙️ Overview

**Automated workflow:**
1. Opens the monthly revenue data source file.  
2. Filters rows from the previous month where `Type = "XX"`.  
3. Copies filtered data to the destination workbook under the `XXX` sheet.  
4. Applies consistent formatting and replicates formulas automatically.  
5. Maps specific codes (column V) to predefined categories (column E).  
6. Saves and closes the workbook safely.  

**Frequency:** Monthly  
**Average runtime:** ~3 minutes  
**Manual time saved:** ≈5.5 hours per year (~0.7 workdays)

---

## 🧩 Technologies Used

| Component | Technology | Purpose |
|------------|-------------|----------|
| **Automation** | `Python 3.10+` | Core automation and data handling |
| **Excel Integration** | `xlwings` | File operations, formatting, and formula replication |
| **Utilities** | `datetime`, `os` | Date filtering and file path management |

---

## 💡 Key Features

- ⚙️ **Automated monthly update** of revenue records  
- 🧾 **Dynamic date filtering** (auto-detects the previous month)  
- 🎨 **Formatting replication** to maintain report consistency  
- 🧮 **Formula propagation** in defined columns (A:D, K)  
- 🧭 **Code mapping** from business keys (column V) to categories (column E)  
- 🧱 **Error-safe and modular structure** for reliability and easy maintenance  

---

## 📂 Repository Structure


---

## 🔄 Detailed Workflow

### 1️⃣ Source Filtering
- Opens the monthly Excel data file.
- Identifies rows where:
  - Column **G = “XX”**  
  - Date in column **AC** corresponds to the previous month.
- Collects all matching rows into a list for transfer.

### 2️⃣ Data Transfer
- Opens the `Revenue_Map.xlsm` file.
- Finds the next empty line in the `XXX` sheet (based on column L).
- Appends all filtered rows sequentially.

### 3️⃣ Formatting Replication
- Copies the cell formatting from the last filled row.  
- Applies it to the newly inserted rows to ensure uniformity.

### 4️⃣ Category Mapping
- Maps column **V** values to column **E** categories using predefined rules:

### 5️⃣ Formula Propagation
- Replicates formulas automatically for new entries in columns:
  - **A:D**
  - **K**
- Ensures all new rows maintain the same calculation logic.

### 6️⃣ Save and Close
- Saves the workbook after successful updates.
- Closes all Excel instances safely.
  
---

## ⭐ Acknowledgements

This automation streamlines the monthly consolidation of commercial revenue data, reducing errors and repetitive work.
It showcases how Python + Excel automation can ensure reliable, auditable, and efficient reporting pipelines in a business environment.
