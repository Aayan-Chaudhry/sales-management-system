# 📊 Sales Management System

A full-stack sales management application built for real-world retail operations, handling sales transactions, customer and product records, custom product formulas, split payments, reporting, label printing workflows, and automated database backups.

## ✨ Features

* Multi-product sales transactions with formula and pricing validation
* Cash, card, e-transfer, custom, and split payment support
* Role-based Admin and Employee access
* Search, filter, edit, and review sales records
* Dashboard tracking daily and all-time sales and revenue
* Reporting for top products, top customers, and ingredient usage
* Automated Excel report generation with OpenPyXL
* Product label queue with print status tracking
* SQLite backups with database integrity checks

## 🛠️ Tech Stack

* **Frontend:** HTML, CSS, JavaScript
* **Backend:** Python, Flask
* **Database:** SQLite
* **Reporting:** OpenPyXL
* **Authentication:** Flask Sessions, role-based access control
* **Automation:** PowerShell, Batch, Python

## 🧠 How It Works

Users authenticate as either an **Employee** or **Admin** before accessing the system.

Employees can record sales, enter multiple products, define product formulas, process different payment methods, manage product labels, and work with current sales records. Transaction data is validated before being saved.

Admins receive additional access to full sales history, transaction deletion, dashboards, reporting, and Excel exports.

Sales records, products, payment data, formulas, and label jobs are stored in SQLite. Related sales and product records use relational database structures, while payment and formula information is stored as structured data within transaction records.

The application also uses SQLite foreign keys and WAL mode, with SQL aggregation and Python processing used to generate business reports.

## 📈 Reporting

The system provides business analytics including:

* Daily and all-time revenue
* Product sales activity
* Top products by revenue
* Top customers by revenue
* Ingredient usage and attributed revenue
* Payment method breakdowns
* Yearly and all-time sales summaries

Reports can also be exported to formatted Excel workbooks using **OpenPyXL**.

## 🏷️ Label Printing

Eligible products can automatically be added to a label printing queue after a sale is recorded.

The queue tracks information such as:

* Product details
* Customer information
* Product formulas and ingredients
* Label templates
* Printed and unprinted status
* Printing errors and retries

Physical label printing uses a local **P-touch printer integration**. Printer-specific configuration and template files are excluded from the public repository.

## 💾 Database & Backups

Application data is stored locally using SQLite.

The included startup utilities can:

* Create timestamped database backups
* Run SQLite integrity checks
* Retain recent backup copies
* Start the Flask application

Local production databases and backup files are excluded from the repository.

## 🚀 Getting Started

### Windows

1. Clone the repository.

```bash
git clone <repository-url>
cd sales-management-system
```

2. Make sure **Python 3** is installed and available on your PATH.

3. Run either launcher:

```text
START.bat
```

or:

```text
START_TRACKER_APP.vbs
```

The launchers install required packages if needed, create a database backup, and start the application.

4. Open:

```text
http://localhost:5000
```

### Manual Setup

Install the required dependencies:

```bash
pip install flask openpyxl
```

Start the application:

```bash
python app.py
```

Then open:

```text
http://localhost:5000
```

## 🔑 Environment Variables

Access credentials and application configuration can be provided using environment variables:

```text
ROUDHA_ADMIN_PIN
ROUDHA_EMPLOYEE_PIN
ROUDHA_TRACKER_SECRET_KEY
```

Local databases, backups, logs, environment files, and printer configuration files are excluded from the repository.

## 🏗️ Project Context

This system was developed to replace manual sales tracking with a centralized application built around real business workflows.

The project focuses on translating operational requirements into software features including transaction management, data validation, access control, automated reporting, reliable data storage, database backups, and integration with physical label-printing hardware.
