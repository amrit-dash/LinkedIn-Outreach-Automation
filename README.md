<div align="center">
  <img src="https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white" alt="Google Apps Script">
  <img src="https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white" alt="LinkedIn">
  <img src="https://img.shields.io/badge/n8n-FF6C37?style=for-the-badge&logo=n8n&logoColor=white" alt="n8n">
  <br/>
  <h1>🚀 LinkedIn Outreach Automation Pro</h1>
  <p><b>A highly robust, automated, and self-hosted LinkedIn Outreach & CRM system built on Google Apps Script and Unipile API.</b></p>
  
  <p>
    <a href="https://github.com/amrit-dash/LinkedIn-Outreach-Automation/issues"><img src="https://img.shields.io/github/issues/amrit-dash/LinkedIn-Outreach-Automation?style=flat-square" alt="Issues"></a>
    <a href="https://github.com/amrit-dash/LinkedIn-Outreach-Automation/network/members"><img src="https://img.shields.io/github/forks/amrit-dash/LinkedIn-Outreach-Automation?style=flat-square" alt="Forks"></a>
    <a href="https://github.com/amrit-dash/LinkedIn-Outreach-Automation/stargazers"><img src="https://img.shields.io/github/stars/amrit-dash/LinkedIn-Outreach-Automation?style=flat-square" alt="Stars"></a>
    <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square" alt="License"></a>
  </p>
</div>

<hr/>

## 📖 Table of Contents
- [Overview](#-overview)
- [Key Features](#-key-features)
- [Architecture & Tech Stack](#-architecture--tech-stack)
- [Repository Structure](#-repository-structure)
- [Prerequisites](#-prerequisites)
- [Installation Guide](#-installation-guide)
- [Configuration](#-configuration)
- [Usage Workflow](#-usage-workflow)
- [Background Workers & Automation](#-background-workers--automation)
- [Troubleshooting & Knowledge Base](#-troubleshooting--knowledge-base)
- [Disclaimer & Best Practices](#-disclaimer--best-practices)
- [License](#-license)

## 🌟 Overview
**LinkedIn Outreach Automation Pro** is an enterprise-grade spreadsheet CRM powered by Google Apps Script. It integrates seamlessly with the [Unipile API](https://www.unipile.com/) to sync LinkedIn accounts, enrich prospect data natively, automate connection requests, and track real-time analytics. 

Whether you are running multi-account outreach campaigns or simple networking automation, this system handles rate limits, webhooks, and background processing—all without leaving Google Sheets!

## ✨ Key Features
- 🔄 **Multi-Account Syncing:** Manage and rotate multiple LinkedIn accounts simultaneously.
- 🎯 **Natively Built Enrichment:** Enrich LinkedIn URLs into detailed prospect profiles directly via the Unipile API.
- 📈 **Campaign Engine:** Bulk-process "Not Started" campaigns into "Active" runs with batch connection requests.
- ⚡ **Background Processing:** Time-driven Google Apps Script triggers handle automated stats gathering and daily sending.
- 🪝 **Webhook Event Handling:** Listen to Unipile webhooks (`reply_received`, `connection_accepted`) directly within Apps Script.
- 📊 **Real-time CRM Analytics:** Tracks sent requests, accepted connections, replied messages, and bounce reasons automatically.

## 🏗 Architecture & Tech Stack
- **Google Apps Script:** Core backend for processing loops, triggers, and the Google Sheets UI.
- **Google Sheets:** The database, CRM interface, and dashboard.
- **Unipile API:** The bridge for LinkedIn actions (invitations, messages, enrichments).
- **n8n (Optional):** Included configurations for visual webhook routing if preferred over Apps Script webhooks.

## 📂 Repository Structure
```text
├── src/                                  # Apps Script Codebase
│   ├── Main.gs                           # UI initialization and entry point
│   ├── Campaigns.gs                      # Core business logic for campaigns
│   ├── Accounts.gs                       # Unipile Account syncing and limits
│   ├── Api.gs                            # Unipile REST API wrappers
│   ├── Prospects.gs                      # Prospect enrichment functionality
│   ├── Webhook.gs                        # Webhook listener logic (doPost)
│   ├── Worker.gs                         # Time-driven trigger functions
│   ├── UI.gs                             # Modal and frontend interactions
│   ├── CampaignForm.html                 # HTML frontend for Campaign Builder
│   └── appsscript.json                   # Google Apps Script Manifest
├── n8n automation/                       
│   └── LinkedIn Outreach Automation.json # Config files for external automation logic
├── .gitignore                            # Excludes sensitive/local files
├── project-plan.md                       # Historical project roadmap
└── README.md                             # This documentation file
```

## 📋 Prerequisites
1. A **Google Workspace / Gmail account** to host the Google Sheet.
2. A **Unipile Account** with a valid API Key and Base URL.
3. Node.js and [clasp](https://github.com/google/clasp) (Command Line Apps Script Projects) for local deployment.

## 🛠 Installation Guide

### 1. Clone the Repository
```bash
git clone https://github.com/amrit-dash/LinkedIn-Outreach-Automation.git
cd LinkedIn-Outreach-Automation
```

### 2. Deploy Code via Clasp
Login to your Google account using `clasp` and push the source files to a Google Sheet Apps Script project.
```bash
npm install -g @google/clasp
clasp login
clasp create --type sheets --title "LinkedIn Automation CRM"
clasp push
```

### 3. Setup the Google Sheet
Ensure your Google Sheet has the following exact tab names:
- `Campaigns`
- `Prospects`
- `Database`
- `Accounts`
- `Invitations`

## ⚙️ Configuration
In your Apps Script editor (`Extensions > Apps Script`), go to **Project Settings > Script Properties** and add:
- `UNIPILE_API_KEY`: Your Unipile API Key.
- `UNIPILE_BASE_URL`: Your Unipile Base URL (e.g., `https://api4.unipile.com:13337/api/v1`).
- `WEBHOOK_MONITORING_ENABLED`: Set to `TRUE` to allow Apps Script to process webhooks.

## 🚀 Usage Workflow
Once the script is deployed, refresh your Google Sheet. You will see a custom menu: **LinkedIn Outreach System**.

1. **🌀 Sync Unipile Accounts:** Validates all tokens and updates the `Accounts` sheet.
2. **🔍 Enrich LinkedIn Prospects:** Processes raw LinkedIn URLs in the `Prospects` sheet to gather Provider IDs.
3. **🚀 Create New Campaign:** Opens a modal to define campaign targets and personalized connection notes.
4. **Start Campaign Now:** 
   - `1 - Create database entries`: Moves enriched prospects into the CRM Database.
   - `2 - Send connection requests`: Dispatches requests via Unipile.
   - `3 - Check connection requests`: Force-verifies accepted connections.

## 🤖 Background Workers & Automation
You don't need to manually click buttons every day. Use the **Automate & Monitor** sub-menu to enable:
- **Campaign Background Worker:** Runs every hour to dispatch pending connection requests.
- **Stats Worker:** Runs periodically to update campaign analytics.
- **Webhook Monitoring:** Automatically processes replies and connections.

## 🚑 Troubleshooting & Knowledge Base
| Issue | Cause | Solution |
| :--- | :--- | :--- |
| **"Daily limit reached"** | Sending limits hit per account. | Configure higher limits in the `Accounts` sheet or wait until tomorrow. |
| **"Failed to enrich"** | Invalid LinkedIn URL or Unipile rate limit. | Verify the URL format. Bulk enrichment has built-in retries. |
| **Missing Webhooks** | n8n / Unipile not pointing to Apps Script. | Deploy Apps Script as a Web App and put the URL into your Unipile Webhook config. |

## ⚠️ Disclaimer & Best Practices
Automating LinkedIn actions violates LinkedIn's Terms of Service. Use this software at your own risk. 
- **Warm up accounts:** Start with 10-20 requests/day.
- **Avoid burst sending:** The script includes intentional `Utilities.sleep()` delays.
- **Monitor bounce rates:** High rejection rates can lead to account restrictions.

## 📄 License
This project is licensed under the [MIT License](LICENSE).
