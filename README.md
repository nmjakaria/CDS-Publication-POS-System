# CDS Publication POS & Inventory Management System

A high-performance, responsive Single-Page Application (SPA) designed exclusively for publication houses to seamlessly manage product listings, inventory increments, vendor supply tracking, sales workflows, customer records, and dynamic financial analysis. 

Built entirely over **Google Web Apps Infrastructure**, this application treats **Google Sheets** as a relational database layer and **Google Apps Script** as a secured backend engine—allowing small to mid-sized teams to deploy an enterprise-grade Point-of-Sale system with zero server maintenance costs.

---

## 🚀 Key Features

### 👤 Role-Based Access Control (RBAC) & Secure Authentication
The system includes an explicitly defined internal permissions engine enforcing layout variations, operational roadblocks, and visibility limits depending on user hierarchy:
* **Admin / Manager:** Has unrestricted structural write/overwrite privileges across all worksheets, master pricing configuration control, dynamic analytics visualization, account deletions, and transaction cancellations.
* **Warehouse Staff:** Authorized access to product records, category grouping tools, vendor inventory manifests, and supply chain input streams.
* **Cashier:** Locked strictly into a dedicated Point of Sale (POS) interface, client profile generation, held checkout summaries, and basic localized invoice lookups.
* **Authentication Flow:** Features an administrative activation switch, automated session logging in an `Activity_Logs` ledger, and background timestamp tracking.

### 📦 Deduplicating Item Stock & Catalog Engine
* **Automated Stock Increments:** When importing bulk files or manually appending inventory metrics, the core database engine checks if a duplicate item code or item name exists under the exact same structural category. Instead of generating redundant database entries, it dynamically increments the piece count (`qty`), updates average ledger rates, and recalculated purchase total values.
* **Multimedia Integration:** Connects seamlessly with standard image files up to 5MB, uploading media arrays straight into a target Google Drive directory (`ItemPOS_Images`), generating open permission view strings, and logging the secure File IDs straight to the data rows.
* **Batch CSV Processing:** Implements a direct custom client-side text processor reading localized `.csv` files natively to stage structural table previews with individual validation flags prior to script commits.

### 💸 Real-time Point of Sale (POS) Interface
* **Dynamic Cart Logic:** A fully reactive frontend built with modular design properties that supports localized instant queries by book names or item codes, automated dynamic calculation thresholds, interactive numerical item pickers, and live cart adjustments.
* **Hold / Resume Workflows:** Allows operators to seamlessly save active checkouts under a `pending` tag to free the register buffer, and safely restore line entries at later intervals without logging dead locks.
* **Multi-Channel Payment Integration:** Natively captures checkout routes through standard cash networks, banking terminals, and regional digital financial networks (`bKash`, `Nagad`).
* **Dual Blueprint Receipt Engine:** Renders on-demand printable assets tailored both for 80mm high-speed standard thermal ticket slips and formatted high-density corporate A4 PDF configurations.

### 📊 Financial Ledger, Analytics & Reporting Matrix
* **Dynamic Supplier & Customer Ledgers:** Seamlessly compiles transactional records to display aggregate purchase metrics, paid totals, and remaining dues via light-grey conditional formatting rules. Includes single-click shortcuts to record partial cash updates.
* **P&L Operational Audits:** Automatically queries transaction boundaries to generate interactive charts tracking Gross/Net Margins, localized expense groups (Transport, Labor, Utilities, Rent), customer ticket value averages, and stock-aging performance distributions.

---

## 🛠️ Tech Stack

* **Runtime Engine:** Google Apps Script (V8 Runtime)
* **Database Tier:** Google Sheets (Relational Architecture with `LockService` concurrency barriers)
* **Storage Framework:** Google Drive API (For asset attachment distribution)
* **Frontend UI Layer:** Modern CSS3 variables with customized theme palettes (`#001f3f` Navy Primary), DataTables framework extension integration, SweetAlert2 wrappers, and FontAwesome 6 icons.
* **Core Core Library:** React (v18 Production Builds via CDN) with raw Babel compilation pipelines.
* **Visualization Engine:** Chart.js (v4 Unified Render Blocks)

---

## 💻 Installation & Setup Guide

Follow this step-by-step layout guide to provision your own sandbox or distribution instance within a personal or corporate Google environment:

### Step 1: Initialize Database Canvas
1. Launch a browser window and create a brand new, empty [Google Sheet](https://sheets.new).
2. Clean your default sheet workspace or keep it as standard. The deployment sequence will override structure programmatically.

### Step 2: Open the Apps Script Editor
1. In the top sheet window toolbar, select **Extensions** > **Apps Script**.
2. Click on the default title (e.g., *Untitled project*) at the upper left and change it to a professional name, such as `CDS_Publication_POS_Engine`.

### Step 3: Populate Script Components
1. Clear any placeholder blocks inside the default workspace file named `code.gs`.
2. Copy the entire functional logic block from your provided `code.js` file and paste it directly into this code interface window.

### Step 4: Run Database Blueprint Scaffolding
1. Inside the file workspace view, look for the function selector drop-down embedded within the control execution toolbar.
2. Select the configuration setup function named **`setup_demo`** (or your designated setup method that generates sheets/tabs).
3. Click the **Run** button (the triangular play icon).
4. **Authorization Request:** An OAuth dialog window will prompt for workspace security consent. Click *Review Permissions*, select your primary account, click *Advanced* (at the base link), and choose *Go to Project Name (unsafe)*. Grant visibility to ensure your script engine can successfully generate sheets, manage directories, and bind parameters.
5. *Result:* The engine will instantly format all structural schema sheets inside your background workbook (`Users`, `Categories`, `Suppliers`, `Purchases`, `Item_Stocks`, `Customers`, `Sales`, `Sale_Items`, `Payments`, `Expenses`, `Import_Logs`, `Settings`, `Activity_Logs`).

### Step 5: Mount Frontend Assets
1. Return to the script editor window sidebar, click the **`+` (Add a file)** icon next to **Files**, and select **HTML**.
2. Type the name exactly as **`index`** (this ensures alignment with the internal Apps Script directive `HtmlService.createHtmlOutputFromFile('index')`).
3. Delete all boilerplate placeholder tags inside the new `index.html` tab.
4. Copy your complete, extensive `index.html` application content containing style classes, React interface nodes, and functional endpoints, then paste it right inside the script window. Save all components with `Ctrl + S`.

### Step 6: Deploy to Live Web App Network
1. Click the **Deploy** button located at the upper-right section of your workspace interface and choose **New deployment**.
2. Click the gear icon next to *Select type* and pick **Web app**.
3. Configure your production parameters precisely as follows:
   * **Description:** `CDS Publication System Production Release v1.0`
   * **Execute as:** `Me (your-email@gmail.com)` *(This gives the web front proxy rights to write updates securely to your database canvas).*
   * **Who has access:** `Anyone` (or restricted to your company domain if executing under an enterprise Google Workspace profile).
4. Click **Deploy**. Copy the generated **Web App URL** from the success popup panel. This provides your live, responsive point of sale environment access link.

---

## 🤝 Contributing

Contributions make the open-source community an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project.
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`).
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`).
4. Push to the Branch (`git push origin feature/AmazingFeature`).
5. Open a Pull Request for administrative code review.

---

## 📄 License

Distributed under the **MIT License**. See standard licensing templates for extended parameter distributions and legal boundaries.

---
## 🎁 Acknowledgments

* Designed and engineered by **Naimullah Md Jakaria**.
* Built using the Google Apps Script & React framework infrastructure.
* Special thanks to AI assistance for collaborating on backend optimization, query caching logic, and documentation.