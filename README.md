# HA Complaints Smart ETL & Self-Referential Pipeline

## 📌 Business Problem
In large-scale home appliance after-sales service operations (managing brands like Gree and EcoStar), raw complaint and dispatch reports export dynamically from ERP systems on a daily basis. 
* **The Challenge:** Operational teams need to track live complaint statuses, categorize complex re-repairs, and add manual/personal notes (`Manual_Remarks`, `My_Personal_Remarks`). However, standard Power Query refreshes overwrite the sheet, wiping out manual data.
* **The Goal:** Build an automated ETL pipeline that ingests the latest pending reports, processes complex operational logic, and **preserves user-entered manual feedback loops** across data refreshes.

---

## 💡 The Solution & Architecture
Developed an advanced **Microsoft Excel Power Query (M Code)** automation pipeline featuring:

1. **Smart Folder Ingestion & Filtering:** 
   * Dynamically targets the designated `HA_Complaints` directory, safely ignoring temporary open system files (`~$`) and automatically sorting/fetching only the latest modified file containing "pending" in its filename.
2. **Automated Status & Keyword Classification Engine:**
   * Implements a robust multi-tier conditional logic to evaluate `JOB_TYPE` and `COMP_WORKING_STATUS`.
   * For **Re-Repairs**, it scans combined system remarks (`REMARKS` + `ACTIVITY_REMARKS`) using keyword matching algorithms (e.g., detecting `leak`, `compressor`, `pcb`, `gas`) to automatically classify issues into **Mobility** or **Field** categories.
3. **Self-Referential Feedback Loop (State Preservation):**
   * Utilizes Excel workbook referencing (`Excel.CurrentWorkbook`) to anchor existing manual remarks. 
   * Performs a **Left Outer Join** on `COMPLAINT_NO` during data refresh, ensuring operational notes and custom remarks never get lost when the core dataset updates.

---

## 🛠️ Tech Stack & Skills Demonstrated
* **Tool:** Microsoft Excel, Power Query
* **Language:** Power Query Formula Language (M Code)
* **Concepts:** Self-Referential ETL Pipelines, Automated File Ingestion, Data Transformation, State Preservation / Feedback Loops, Conditional Business Logic.

---

## 📂 Repository Structure
* `automation_master_script.pq` - Contains the full, production-ready Power Query M-code.
* `README.md` - Project documentation.

---

## 🚀 How to Implement
1. Clone this repository.
2. Save your raw ERP export files inside a local folder path matching the script reference:
   `C:\Users\PC\Desktop\HA_Complaints`
3. Open Excel, launch the **Power Query Advanced Editor**, paste the contents of `automation_master_script.pq`, and link your workbook name to `HA_Complaints`.