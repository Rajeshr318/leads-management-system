# Leads Management System (LMS)

A secure, full-stack Leads Management System featuring an interactive 2D spreadsheet data grid, role-based authentication, and segregated workflows for Administrators and Counselors.

## Features
* **Role-Based Access Control (RBAC):** Secure login gateway separating full administrative governance from restricted counselor operational views.
* **Row-Level Security (RLS):** Automatically scopes and filters spreadsheet rows so counselors see only their assigned leads.
* **Interactive 2D Data Grid:** Spreadsheet matrix interface supporting dynamic Excel/CSV ingestion, formula bar tracking ($fx$), column letters, and date filters (`DD-MM-YYYY`).
* **Batch Operations:** Multi-row selection with batch deletion controls restricted to authorized admin users.

## Tech Stack
* **Frontend:** React.js, Vite, HTML5, CSS3
* **Backend Architecture:** Python, FastAPI, SQLAlchemy
* **Database & Security:** PostgreSQL, JWT Authentication, RBAC

## Getting Started
1. Clone the repository:
   ```bash
   git clone [https://github.com/Rajeshr318/leads-management-system.git](https://github.com/Rajeshr318/leads-management-system.git)
