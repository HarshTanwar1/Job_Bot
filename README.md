<div align="center">

# 🤖 Job_Bot

An end-to-end RPA bot that scrapes, structures, filters & emails job-market data — fully hands-free

_Driving Chrome across five major job portals to turn hours of manual research into a one-click, automated report._

<br>

![UiPath](https://img.shields.io/badge/UiPath-Studio_21.10-FA4616?style=for-the-badge&logo=uipath&logoColor=white)
![RPA](https://img.shields.io/badge/RPA-Automation-0F4C81?style=for-the-badge)
![.NET](https://img.shields.io/badge/.NET-VB.NET-512BD4?style=for-the-badge&logo=dotnet&logoColor=white)
![Excel](https://img.shields.io/badge/Excel-Automation-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![SMTP](https://img.shields.io/badge/Email-SMTP-D14836?style=for-the-badge&logo=gmail&logoColor=white)

</div>

<br>

## 🎯 The Problem It Solves

Searching job portals by hand is slow, repetitive, and error-prone — open a site, run a search,
click each listing, copy the salary, benefits, and requirements, paste them into a spreadsheet,
repeat. **Job_Bot does all of it automatically.**

It drives a real Chrome browser through **five popular job sites — LinkedIn, Indeed, Glassdoor,
Monster, and SimplyHired** — reads each posting, extracts the details that matter, lets you filter
the results by keyword, exports everything to Excel, and emails you the finished report. No manual
copy-pasting, no missed listings.

<br>

## 🌟 Highlights

|                                 |                                                                                                                                        |
| ------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------- |
| 🔎 **5 job portals automated**  | LinkedIn · Indeed · Glassdoor · Monster · SimplyHired                                                                                  |
| 🧩 **Structured extraction**    | Parses raw postings into 7 clean fields (compensation, benefits, experience, responsibilities, qualifications, schedule, requirements) |
| 🎛️ **Interactive filtering**    | Filter results by any column + keyword at runtime — no code edits                                                                      |
| 📊 **Excel reporting**          | Auto-writes results and filtered copies to spreadsheets                                                                                |
| 📧 **Automated email delivery** | Sends the final report over SMTP to any recipient                                                                                      |
| ♻️ **Modular & reusable**       | Shared sub-workflows for extraction & filtering, reused across every scraper                                                           |
| 🛡️ **Resilient**                | Try/Catch + app-state checks survive slow loads, pop-ups & layout quirks                                                               |

<br>

## ✨ Features & Functionalities

- **Multi-portal job scraping** — Dedicated workflows scrape listings from LinkedIn, Indeed,
  Glassdoor, Monster, and SimplyHired.
- **Structured data extraction** — A reusable `Data_Extraction` workflow parses each raw job-detail
  blob into clean fields: _compensation, benefits, experience, responsibilities, qualifications,
  schedule,_ and _requirements_.
- **Personal & organization data collection** — `Personal.xaml` and `Organization.xaml` gather
  employee-level and company-level records (name, address, job title, industry, company name,
  employment status, and more).
- **Excel export** — Collected data is written to workbooks (`Employee_Data.xlsx`,
  `Employment.xlsx`) via Write Range.
- **Interactive keyword filtering** — `Data_Filtering` prompts the user to choose a filter column
  and enter a keyword, then removes non-matching rows and saves a filtered copy
  (`Filtered_Employee_Data.xlsx`, `Filtered_Employment.xlsx`).
- **Automated email delivery** — Results are emailed to a user-supplied address via SMTP.
- **Resilient navigation** — Try/Catch blocks and `Check App State` / _Target appears_ checks
  handle pages that load slowly or differ between listings.

<br>

## 🛠️ Tech Stack

| Area                    | Technology                                                                                           |
| ----------------------- | ---------------------------------------------------------------------------------------------------- |
| **Automation platform** | UiPath Studio 21.10 (Workflow / `.xaml`)                                                             |
| **Runtime**             | .NET Framework (Legacy profile)                                                                      |
| **Expression language** | VB.NET                                                                                               |
| **Web / UI automation** | `UiPath.UIAutomation.Activities` — Chrome automation, selectors, Extract Table Data, Check App State |
| **Spreadsheets**        | `UiPath.Excel.Activities` — Read/Write Range                                                         |
| **Email**               | `UiPath.Mail.Activities` — Send SMTP Mail Message                                                    |
| **Core logic**          | `UiPath.System.Activities` — control flow, assigns, input dialogs                                    |
| **Version control**     | Git / GitHub                                                                                         |

<br>

## 🏗️ How It's Built

```
Main.xaml                      # Entry point (workflow shell)
│
├── Personal.xaml              # Orchestrates job scraping + employee data + filtering + email
│   ├── Extract_Data_Linkedin.xaml      ┐
│   ├── Extract_Data_Indeed.xaml        │
│   ├── Extract_Data_Glassdoor.xaml     ├─ Site scrapers → each invokes Data_Extraction
│   ├── Extract_Data_Monster.xaml       │
│   └── Extract_Data_SimplyHired.xaml   ┘
│
├── Organization.xaml          # Collects organization/company data + filtering + email
│
├── Data_Extraction.xaml       # ♻️ Reusable: parses raw job text → structured fields
└── Data_Filtering.xaml        # ♻️ Reusable: keyword-based row filtering + Excel write
```

> **Design principle:** common logic lives in shared, argument-driven sub-workflows
> (`Data_Extraction`, `Data_Filtering`) instead of being copy-pasted into every scraper — keeping
> the project DRY and easy to extend with new job sites

<br>

## 🚀 Getting Started

### Prerequisites

- **Windows** OS
- **[UiPath Studio](https://www.uipath.com/product/studio)** `21.10.x` (Community Edition is free)
- **Google Chrome** with the **UiPath Browser Extension** enabled
  (_Studio → Tools → UiPath Extensions → Chrome_)
- A reachable **SMTP server / email account** for the email step

### Run It

1. **Clone the repository**

   ```bash
   git clone https://github.com/HarshTanwar1/Job_Bot.git
   ```

2. **Open in UiPath Studio**
   - Launch Studio → **Open a Local Project** → select `project.json`.
   - Studio restores dependencies automatically (Excel, Mail, System, UIAutomation). If not,
     restore them via **Manage Packages**.

3. **Verify the Chrome extension** is installed and enabled (the bot uses _Use Browser Chrome_).

4. **Configure email** — Open `Personal.xaml` / `Organization.xaml` and set your SMTP server, port,
   and credentials in the **Send SMTP Mail Message** activity.

5. **Run a workflow**
   - **Employee + job data:** open and **Run** `Personal.xaml`
   - **Organization data:** open and **Run** `Organization.xaml`
   - Follow the on-screen dialogs to enter your email, choose whether to filter, pick a column, and
     enter a keyword.

6. **Check the output** — Results land in the `.xlsx` files (`Employee_Data.xlsx`,
   `Employment.xlsx`, and the `Filtered_*.xlsx` copies) **and** in your inbox.

> ⚠️ **Note:** Scrapers depend on each site's current page layout. If a site changed its markup
> since this project was built, some selectors may need to be re-indicated in UiPath Studio

<br>

## 📚 What I Learned

- Building **modular, reusable UiPath workflows** by extracting shared logic into separate
  argument-driven `.xaml` files instead of duplicating it across every scraper.
- Designing robust **web selectors** and handling the reality that every job portal — and even
  individual listings on the same site — renders its pages differently.
- Using **`Check App State` / _Target appears_** and **Try/Catch** to make automation resilient to
  slow loads, missing elements, and pop-ups instead of crashing on the first surprise.
- Passing data between workflows with **In / InOut arguments** and moving structured data around
  using **DataTables**.
- Integrating **Excel read/write** and **SMTP email** to turn raw scraped data into a finished,
  deliverable report.
- Adding **interactive input dialogs** so the same bot can be re-run with different keywords,
  filter criteria, and recipients without touching the workflow.

<br>

## 🔮 Future Improvements

- **Stronger error handling & logging** — Structured logs, retry scopes, and failure screenshots.
- **Selector resilience** — Anchor-based / fuzzy selectors to reduce maintenance as sites change.
- **Deduplication & validation** — Drop duplicate listings and validate fields before export.
- **Scheduling** — Run unattended via UiPath Orchestrator/Assistant for automated daily digests.
- **Auth handling** — Gracefully handle login walls (e.g. LinkedIn's auth wall) that block guest
  scraping.

<br>

---

<div align="center">

⭐ _If you found this project helpful or interesting, consider giving it a star!_ ⭐

</div>
