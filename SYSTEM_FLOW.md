# 🔄 System Flow Documentation

This document explains the end-to-end data flow, UI interactions, and background processes of the **LinkedIn Outreach Automation** system.

---

## 1. Setup & Initialization

### A. Credentials Configuration
- The system stores credentials in the **Credentials** sheet.
- **Row 1, Col B:** `apiKey` (Unipile API Key)
- **Row 2, Col B:** `dsn` (Unipile DSN URL)
- Every time an action is executed, `getCredentials()` (in `Api.gs`) reads these cells to construct the API authorization headers.

### B. UI Initialization (`Main.gs`)
- When the spreadsheet opens, the `onOpen()` trigger runs.
- It builds a custom UI menu: **LinkedIn Outreach System**.
- This menu exposes actions like Syncing Accounts, Enriching Prospects, Creating Campaigns, starting Campaign processes, and managing Automation background workers.
- The `updateGlobalStats()` function is called silently to refresh the dashboard numbers on open.

---

## 2. Unipile Accounts Syncing

### A. Manual Trigger (`Accounts.gs`)
- **Action:** User clicks **LinkedIn Outreach System > 🌀 . Sync Unipile Accounts**.
- **Process:**
  1. Script fetches connected accounts from Unipile via `GET /api/v1/accounts`.
  2. Compares the retrieved accounts with those listed in the **Accounts** sheet.
  3. Updates connection statuses (e.g., marks as `Active` if the Unipile source status is `OK`).
  4. If new accounts are found, they are appended as new rows.
  5. Any accounts listed in the sheet but *not* returned by Unipile are marked as `Inactive`.

---

## 3. Prospect Management & Native Enrichment

### A. Gathering Raw Prospects
- The user pastes raw prospect data into the **Prospects** sheet.
- Column 3 (Index C) must contain the prospect's LinkedIn URL.

### B. Enrichment Process (`Prospects.gs`)
- **Action:** User clicks **LinkedIn Outreach System > 🔍 . Enrich LinkedIn Prospects**.
- **Process:**
  1. Prompts the user to select an active Unipile sending account.
  2. Extracts the public LinkedIn ID from the URL (`extractLinkedInId`).
  3. Calls Unipile API: `GET /api/v1/users/{identifier}?account_id={accountId}`.
  4. **Data Extraction:** Parses First Name, Last Name, Country, City, Headline, Company Name, and Website.
  5. **Crucial Step:** Retrieves the unique `provider_id` from Unipile and stores it in the sheet.
  6. Sets Column L (Index 11) to `Yes` indicating successful enrichment. 

---

## 4. Campaign Execution Workflow

### A. Campaign Creation (`UI.gs` & `CampaignForm.html`)
- User clicks **🚀 . Create New Campaign**.
- An HTML modal opens, collecting:
  - Campaign Name, Number of Prospects.
  - Connection Note (supports `$name` variable).
  - Messages 1, 2, and 3 along with hour delays.
- Upon submission, a new row is appended to the **Campaigns** sheet with status `Not Started`.

### B. Step 1: Create Database Entries (`Campaigns.gs`)
- **Action:** User clicks **1 - Create database entries**.
- **Process:**
  1. Identifies campaigns with `Not Started` status.
  2. Selects `n` prospects from the **Prospects** sheet based on target count.
  3. Ensures those prospects are enriched (inline enrichment happens if they aren't).
  4. Moves the successfully enriched prospects into the **Database** sheet.
  5. Sets the Connection Request Status to `Pending`.
  6. Updates the Campaign status to `Active`.

### C. Step 2: Sending Connection Requests (`Campaigns.gs`)
- **Action:** User clicks **2 - Send connection requests**.
- **Process:**
  1. Locates all database rows for the campaign where Status is `Pending`.
  2. Verifies account limits (Daily limits defined in **Accounts** sheet).
  3. Sends connection invitations via Unipile API: `POST /api/v1/users/invite` with the Connection Note payload.
  4. Updates the DB row Status to `Sent` and records the Timestamp.
  5. The `invitation_id` returned by Unipile is logged into the hidden **Invitations** sheet for future tracking (e.g., withdrawing invites).

### D. Step 3: Force Checking Requests (`Campaigns.gs`)
- **Action:** User clicks **3 - Check connection requests**.
- **Process:** Manually iterates over `Sent` prospects, checks their profile via Unipile `GET` to see if `connected_at` is populated. If yes, marks as `Accepted`.

---

## 5. Webhook Integration & Event Monitoring

The system monitors two primary events: `new_relation` (connection accepted) and `message_received` (prospect replied).
There are two ways this is handled depending on user configuration:

### Option A: Apps Script doPost Handler (`Webhook.gs`)
- The Unipile webhook points directly to the Google Apps Script Web App URL.
- **Payload Parsing:**
  - If `event == 'new_relation'`: Finds the prospect in the **Database** sheet by matching `account_id` and `user_provider_id`. Sets `connection_request_status` to `Accepted`, checks the boolean box, and logs the time.
  - If `event == 'message_received'`: Finds prospect via `sender.attendee_provider_id`. Sets `reply_received` to TRUE, logs `reply_text` and `reply_time`. Skips logging if the message was sent by the user themselves.

### Option B: n8n Workflow Automation
- Handled via `n8n automation/LinkedIn Outreach Automation.json`.
- **Node 1 (Webhook):** Listens for Unipile POST payloads.
- **Node 2 (Switch Router):** Routes the flow based on `event == 'message_received'` or `event == 'new_relation'`.
- **Node 3 (Get Database Entry):** Searches the Google Sheet Database by `sending_account` and `provider_id`.
- **Node 4 (Filter):** Ensures a valid `row_number` is found.
- **Node 5 (Update Sheet):**
  - For Connections: Updates `connection_accepted` to TRUE and logs the time.
  - For Replies: Updates `reply_received` to TRUE and logs the message text and time.

---

## 6. Background Automation Workers (`Worker.gs`)

### A. The Campaign Process Worker (`processCampaignsWorker`)
- Triggered by Apps Script Time-Driven trigger (e.g., every 15 minutes).
- **Core Responsibilities:**
  1. **7-Day Invite Cleanup:** Looks for requests marked `Sent` older than 7 days that haven't been accepted. Finds the `invitation_id` in the **Invitations** sheet and calls `DELETE /api/v1/users/invite/sent/{id}`.
  2. **Follow-Up Messaging:** Looks for prospects marked `Accepted`.
     - If `msg1_status` is Pending -> Sends Message 1.
     - If `msg1_status` is Sent -> Checks if time elapsed > `delay2` hours -> Sends Message 2.
     - If `msg2_status` is Sent -> Checks if time elapsed > `delay3` hours -> Sends Message 3.
  3. **Stop on Reply:** If the DB indicates `reply_received` == TRUE, it skips that prospect entirely, preventing automated follow-ups to people who have replied.

### B. The Statistics Worker (`processStatsWorker` / `updateGlobalStats`)
- Triggered every 10 minutes or on every manual sheet edit (`onEdit`).
- **Process:**
  1. Reads the entire **Database** sheet.
  2. Rolls up numbers per Campaign: Sent requests, Accepted connections, Messages sent, Replies received.
  3. Rolls up numbers per Account: Tracks "Sent Today" metrics based on current timestamps to enforce daily rate limits.
  4. Writes these aggregated numbers directly to the `Campaigns` and `Accounts` sheets to maintain real-time dashboards.
