# Report Email to FSO – Outlook Web Add-in

Modern web add-in that adds a **Report Email to FSO** ribbon button.  
Works in **New Outlook for Windows**, Outlook on the web, classic Outlook, and Outlook on Mac.

When you click the button it opens a new email to `jrviar@tracewell.com` with the original message attached.

---

## Deploy with GitHub Pages (no local server needed)

### Step 1 – Create the GitHub repository

1. Go to https://github.com/new
2. Repository name: **`outlook-fso-report`** (recommended – matches the URLs already in the manifest)
3. Make it **Public** (required for free GitHub Pages)
4. Do **not** add a README, .gitignore, or license (we already have the files)
5. Click **Create repository**

### Step 2 – Upload these files

You can do this in the browser or with Git.

**Browser method (easiest):**

1. On the new empty repo page click **uploading an existing file**
2. Drag **everything** inside the `OutlookAddIn_FSO_Report_Web` folder into the browser:
   - `manifest.xml`
   - `README.md`
   - `.nojekyll`
   - `package.json` (optional)
   - `assets/` folder
   - `src/` folder
3. Commit the files

**Or with Git (if you prefer the command line):**

```bash
cd OutlookAddIn_FSO_Report_Web
git init
git add .
git commit -m "Initial Outlook FSO Report add-in"
git branch -M main
git remote add origin https://github.com/jrviar2-twell/outlook-fso-report.git
git push -u origin main
```

### Step 3 – Turn on GitHub Pages

1. In the repo go to **Settings → Pages**
2. Under **Source** choose **Deploy from a branch**
3. Branch: **main** / folder: **/ (root)**
4. Click **Save**
5. Wait 1–2 minutes. Your site will be live at:

   ```
   https://jrviar2-twell.github.io/outlook-fso-report/
   ```

### Step 4 – Put your real username into the manifest

Open `manifest.xml` and replace **every** occurrence of:

```
jrviar2-twell
```

with your actual GitHub username (example: if your username is `jsmith`, the base URL becomes `https://jsmith.github.io/outlook-fso-report`).

You can do a find-and-replace in any text editor, or tell me your GitHub username and I will update the file for you.

After changing the file, commit and push again (or re-upload the updated `manifest.xml`).

### Step 5 – Sideload the add-in in New Outlook

1. Open **New Outlook** (or go to https://outlook.office.com)
2. Go to https://aka.ms/olksideload  
   **or** Settings (gear) → General → Manage add-ins → **Add from file**
3. Choose **Upload custom add-in** → **Upload manifest file**
4. Select the `manifest.xml` that now contains your real GitHub username
5. The add-in should appear. Open any email and look for the **FSO Report** group on the ribbon.

---

## What the button does

| Field        | Value                                      |
|--------------|--------------------------------------------|
| To           | jrviar@tracewell.com                       |
| Subject      | Reported Email: <original subject>         |
| Body         | Short report with original from / time     |
| Attachment   | The original email (Outlook item)          |

You review the message and click **Send**.

---

## Changing the report address later

Edit `src/commands/commands.js` and change the line:

```js
toRecipients: ["jrviar@tracewell.com"],
```

Commit, push, and the change is live within a minute or two (GitHub Pages cache is short).

---

## Project layout

```
outlook-fso-report/          ← root of the GitHub repo
├── .nojekyll                ← tells GitHub not to use Jekyll
├── manifest.xml             ← the file you sideload
├── README.md
├── assets/                  ← icons
└── src/
    └── commands/
        ├── commands.html
        └── commands.js
```

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Button does not appear | Make sure you uploaded the **updated** manifest (with your real username). Restart New Outlook. |
| "Add-in error" or blank | Open the GitHub Pages URL in a browser and confirm the files load (e.g. …/src/commands/commands.html). |
| Icons missing | Confirm `assets/icon-32.png` etc. are reachable at the GitHub Pages URL. |
| Still seeing localhost | You forgot to replace the placeholder username in the manifest. |

---

Once you tell me your **GitHub username**, I can update `manifest.xml` with the real URLs for you so you only have to upload and turn on Pages.
