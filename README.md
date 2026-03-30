# Neodrift Creative Manager

A production-ready **Google Apps Script web app** for managing ad creatives across Google Ads, Meta Ads, Amazon Ads, and other platforms. Built for D2C brands running multi-platform advertising campaigns.

Upload, organize, review, approve, and track ad creatives — all backed by a Google Sheet and Google Drive, with optional **AI-powered naming** via the Claude API.

---

## Features

### Core Creative Management
- **Upload** images and videos (drag-and-drop, supports resumable uploads for large files)
- **Gallery views** — Grid, List, and Folder view (organized by product category)
- **Bulk operations** — Select multiple creatives for category assignment, AI naming, or deletion
- **Approval workflow** — Pending → Approved (with star rating) → Tested on Ads
- **Changes Required** flow with rejection notes and file replacement
- **Revert to Pending** — Admin can revert any approved/tested/rejected creative back to pending

### AI-Powered Naming (Claude API)
- Analyzes creative thumbnails to identify products and suggest standardized filenames
- Matches against your product catalog automatically
- Bulk AI naming for selected creatives
- Password-protected (time-based IST code) to control API usage

### Organization
- **Product categories** with auto-folder creation in Google Drive
- **Custom attributes** — Create dropdown tags (e.g., "Type: Ad/Lifestyle/Infographic") and filter by them
- **Drive sync** — Two-way sync between your Sheet and Drive folder
- **Import** — Import creatives from any Google Drive folder by search or link

### Ad Testing Tracker
- Track which creatives have been tested on **Google Ads**, **Meta Ads**, **Amazon Ads**, and **Others**
- Custom label for "Others" platform (e.g., Flipkart, Snapchat)
- Status auto-updates to "Tested on Ads" when any platform is checked

### Activity Log & Undo
- Full activity log tracking all actions (uploads, approvals, rejections, renames, deletes, bulk ops)
- **Password-protected Undo** — Revert actions directly from the activity log
- Timestamps, old/new values, and user attribution for every action

### Integrity & Error Checking
- Detect duplicate codes, duplicate names, and missing Drive files
- **One-click fix buttons** for each error type
- Sheet ↔ Drive sync status check

### Multi-User Support
- Role-based access: **Admin** (full access) and **Designer** (upload + view)
- User management in Settings
- Session persistence (stay logged in)

### Responsive Design
- Full mobile support with bottom navigation bar
- Touch-friendly swipe navigation in creative preview
- Image zoom on tap

---

## File Structure

```
neodrift-creative-manager/
├── Code.gs              # Server-side Google Apps Script (all backend logic)
├── index.html           # Client-side UI (single-file HTML/CSS/JS)
├── appsscript.json      # Apps Script manifest (permissions, timezone, runtime)
└── README.md            # This file
```

---

## Deployment Guide

### Prerequisites
- A Google account
- A Google Sheet (new or existing)

### Step 1 — Create the Apps Script Project

1. Go to [Google Sheets](https://sheets.google.com) and create a new spreadsheet
2. Name it something like `Neodrift Creative Manager`
3. Go to **Extensions → Apps Script**
4. This opens the Apps Script editor

### Step 2 — Add the Code

**Option A: Manual Copy-Paste**

1. In the Apps Script editor, replace the contents of `Code.gs` with the contents of [Code.gs](./Code.gs)
2. Click **+ (Add file) → HTML** and name it `index` (it will become `index.html`)
3. Paste the contents of [index.html](./index.html) into this file
4. Click the gear icon ⚙️ (Project Settings) in the left sidebar
5. Check **"Show appsscript.json manifest file in editor"**
6. Open `appsscript.json` and replace its contents with [appsscript.json](./appsscript.json)

**Option B: Using clasp (CLI)**

```bash
# Install clasp globally
npm install -g @google/clasp

# Login to your Google account
clasp login

# Clone your Apps Script project (get the script ID from the URL)
clasp clone <SCRIPT_ID>

# Copy the files into the cloned directory
cp Code.gs index.html appsscript.json ./

# Push to Apps Script
clasp push
```

### Step 3 — Deploy as Web App

1. In the Apps Script editor, click **Deploy → New deployment**
2. Click the gear icon ⚙️ next to "Select type" and choose **Web app**
3. Configure:
   - **Description**: `Creative Manager v7`
   - **Execute as**: `User accessing the web app`
   - **Who has access**: Choose based on your needs:
     - `Only myself` — Personal use
     - `Anyone within [your org]` — Team use (Google Workspace)
     - `Anyone` — Public access (requires Google login)
4. Click **Deploy**
5. **Authorize** the required permissions when prompted
6. Copy the **Web app URL** — this is your app!

### Step 4 — Initialize

1. Open the web app URL in your browser
2. Login with the default credentials: `admin` / `admin123`
3. Go to **Settings** and:
   - **Change the admin password** (edit the Users sheet directly or add a new admin user and delete the default)
   - **Add your Claude API key** (optional, for AI naming — get one from [Anthropic Console](https://console.anthropic.com))
   - **Add your products** to the product catalog
   - **Create custom attributes** if needed (e.g., "Type" with options "Ad, Lifestyle, Infographic")

---

## Google Sheet Structure

The app automatically creates and manages these sheets:

| Sheet | Purpose |
|-------|---------|
| `Creatives` | Main data — one row per creative with code, name, status, file ID, ratings, test flags, etc. |
| `Users` | Login credentials and roles |
| `Settings` | App configuration (API key, Drive folder ID, nomenclature format) |
| `Products` | Product catalog for categorization |
| `Attributes` | Custom dropdown attributes |
| `ActivityLog` | Full audit trail of all actions |

---

## AI Naming Setup

The AI naming feature uses the **Claude API** (Anthropic) to analyze creative thumbnails and suggest standardized filenames.

### How It Works
1. Fetches a thumbnail of the creative from Google Drive
2. Sends it to Claude Haiku with your product catalog
3. Claude identifies the product and suggests a filename in the format: `{product}_{platform}_{variant}_{date}`
4. The name is applied and the creative is moved to the matching product folder

### Setup
1. Get an API key from [console.anthropic.com](https://console.anthropic.com)
2. Go to **Settings → Claude API** in the app and paste your key
3. Add products to the catalog (Settings → Products)
4. Use the 🤖 button on any creative or bulk-select and click "AI Name"

### Access Control
AI naming is protected by a time-based password to prevent accidental API charges:
- Password format: `DDMMYYHH` in IST (e.g., `31032614` for March 31, 2026 at 2 PM IST)
- Unlocks for 15 minutes per session

---

## Usage Tips

### Bulk Workflow
1. Go to **Creatives** tab
2. Click **☑ Select** to enter selection mode
3. Select creatives (or use folder view → "Select All" per folder)
4. Use the bulk bar: assign category, AI name, or delete

### Approval Workflow
1. Designers upload creatives → status: **Pending Approval**
2. Admin opens each in **Pending** tab → rates with stars → **Approve** or **Reject**
3. Rejected creatives get notes → designer replaces file → resubmits
4. Approved creatives move to **Testing** tab for ad platform tracking

### Import Existing Creatives
1. Go to **Creatives → 📥 Import**
2. Search for a Drive folder or paste a link
3. Files are **copied** (not moved) into your managed Drive folder
4. Each imported file gets a unique code and "Pending Approval" status

---

## Customization

### Nomenclature Format
In Settings, you can customize the AI naming format. Default: `{product}_{platform}_{variant}_{date}`

### Custom Attributes
Create any dropdown attribute in Settings → Custom Attributes. Examples:
- `Type`: Ad, Lifestyle, Infographic, A+ Content
- `Campaign`: Diwali Sale, Summer Launch, BAU
- `Designer`: Person A, Person B

These appear in the creative detail modal and as gallery filters.

### Web App Access
To change who can access the app after deployment:
1. Go to Apps Script → **Deploy → Manage deployments**
2. Edit the deployment and change the access level

---

## Updating the App

When you update the code:

1. Make your changes in the Apps Script editor
2. Go to **Deploy → Manage deployments**
3. Click the ✏️ edit icon on your deployment
4. Change **Version** to "New version"
5. Click **Deploy**

The URL stays the same — users just refresh the page.

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Not authorized" error | Re-deploy and re-authorize permissions |
| Thumbnails not loading | Check that files have "Anyone with link" sharing |
| AI naming fails | Verify API key in Settings; check Anthropic console for billing |
| Sync shows mismatch | Click 🔄 Sync on the Creatives tab to reconcile |
| Slow on large datasets | The app handles 500+ creatives well; for 1000+, consider archiving old creatives |

---

## Tech Stack

- **Backend**: Google Apps Script (V8 runtime)
- **Frontend**: Vanilla HTML/CSS/JS (no frameworks, no build step)
- **Storage**: Google Sheets (database) + Google Drive (file storage)
- **AI**: Claude Haiku API via Anthropic (optional)
- **Auth**: Simple username/password stored in the Users sheet

---

## License

MIT — Free to use, modify, and distribute.

---

## Credits

Built for [NEODRIFT](https://neodrift.in) — Premium car and bike accessories on Amazon India.

Part of the [Ready-To-Use Business Web Apps](https://github.com/your-username/Ready-To-Use-Business-Web-Apps) collection.
