# ScholarKit

A personal-use academic writing toolkit for Microsoft Word and Google Docs. Four tools in one pane:

1. **Phrases**: 476 original academic phrase templates, organized by paper section (Introduction, Methods, Discussion...) and writing function (hedging, comparing, defining...). Click to insert.
2. **Rewrite**: AI paraphrasing with 3 options per pass (academic, formal, simplify, concise) plus academic word choice suggestions. Powered by your own free Gemini API key.
3. **Similarity**: checks your draft against your own reference PDFs and documents for overlapping passages of 5+ identical words. Catches accidental copy-paste before your professor does.
4. **Settings**: where your API key lives.

Cost to run: $0. Gemini free tier covers personal use. Hosting is free GitHub Pages.

---

## What is in this folder

```
scholarkit/
  word-addin/
    manifest.xml        <- the file you give to Word
    taskpane.html       <- the pane UI
    taskpane.css
    taskpane.js
    phrasebank.json     <- the phrase database (shared by both platforms)
    assets/             <- icons
  docs-addon/
    Code.gs             <- paste into Apps Script
    Sidebar.html        <- paste into Apps Script
  README.md             <- this file
```

---

## Step 1: Get a free Gemini API key (5 minutes, do this once)

1. Open a web browser and go to **aistudio.google.com**
2. Sign in with your Google account
3. Look for a button that says **Get API key** (top left area). Click it.
4. Click **Create API key**
5. A long string of letters and numbers appears. Click the **copy** icon next to it.
6. Paste it somewhere safe for now, like a private note. You will paste it into ScholarKit's Settings tab later.

Treat this key like a password. Do not put it in any public file or repo.

---

## Step 2: Put the Word add-in online (10 minutes, do this once)

Word add-ins must load from a website with https. GitHub Pages does this for free.

**Important: name the repo exactly `scholarkit` (all lowercase).** Every URL in the project assumes `https://raiderj77.github.io/scholarkit/`. If you use a different name, see Troubleshooting at the bottom.

1. Go to **github.com** and sign in
2. Click the **+** in the top right, then **New repository**
3. Repository name: type `scholarkit`
4. Set it to **Public** (Pages on private repos needs a paid plan)
5. Do NOT put your API key anywhere in this repo. There is no key in any of these files, keep it that way.
6. Click **Create repository**
7. On the new empty repo page, click the link that says **uploading an existing file**
8. On your computer, open the scholarkit folder. Drag everything inside it (the `word-addin` folder, `docs-addon` folder, and README.md) into the upload box on GitHub
9. Wait for the files to finish uploading, then click the green **Commit changes** button
10. Now turn on Pages: click **Settings** (top of the repo), then **Pages** in the left sidebar
11. Under "Build and deployment", set Source to **Deploy from a branch**
12. Set Branch to **main** and folder to **/ (root)**, then click **Save**
13. Wait 2 to 3 minutes. Refresh the page. You will see "Your site is live at https://raiderj77.github.io/scholarkit/"
14. Test it: open **https://raiderj77.github.io/scholarkit/word-addin/taskpane.html** in your browser. You should see the ScholarKit pane. If you get a 404, wait another minute and refresh.

---

## Step 3: Load the add-in into Word

First, download `manifest.xml` to your computer if it is not already there (it is in the `word-addin` folder of this bundle).

### Option A: Word on the web (easiest, works everywhere)

1. Go to **office.com**, sign in, and open any Word document (or create a new one)
2. On the **Home** ribbon, click **Add-ins** (a puzzle-piece or grid icon). If you do not see it, look on the **Insert** ribbon instead.
3. Click **More Add-ins** (or **Get Add-ins**)
4. In the window that opens, click the **MY ADD-INS** tab at the top
5. Click **Upload My Add-in** (it may be a small link in the corner)
6. Click **Browse**, pick your `manifest.xml` file, click **Upload**
7. A ScholarKit button appears on the Home ribbon. Click it. The pane opens on the right.

Note: with this method, the add-in stays attached to documents you use it in. For a new document, repeat the upload (it takes 10 seconds once you have done it).

### Option B: Word desktop on Windows (sticks permanently)

1. Make a folder on your C drive: open File Explorer, go to **C:\\**, right-click, **New > Folder**, name it `addins`
2. Copy `manifest.xml` into `C:\addins`
3. Share the folder with yourself: right-click the `addins` folder, click **Properties**, click the **Sharing** tab, click **Share...**
4. In the dropdown, pick your own username, click **Add**, then click **Share**, then **Done**
5. Still in Properties > Sharing, note the **Network Path**. It looks like `\\YOUR-PC-NAME\addins`. Write it down exactly.
6. Open Word. Click **File > Options > Trust Center > Trust Center Settings... > Trusted Add-in Catalogs**
7. In the **Catalog Url** box, paste your network path from step 5, click **Add catalog**
8. Check the **Show in Menu** box next to the new entry, click **OK**, click **OK** again
9. Close Word completely and reopen it
10. Click **Home > Add-ins > More Add-ins** (older Word: **Insert > My Add-ins**), then click the **SHARED FOLDER** tab
11. Click **ScholarKit**, then **Add**. The button now lives on your Home ribbon in every document.

### Option C: Word desktop on Mac

1. In Finder, press **Cmd+Shift+G** and paste this path:
   `/Users/YOURUSERNAME/Library/Containers/com.microsoft.Word/Data/Documents/wef`
   (replace YOURUSERNAME with your Mac username; create the `wef` folder if it does not exist)
2. Copy `manifest.xml` into that folder
3. Restart Word. Go to **Insert > Add-ins > My Add-ins** and ScholarKit appears under Developer Add-ins.

### First run

Open the pane, click the **Settings** tab, paste your Gemini key, leave the model as `gemini-2.5-flash`, click **Save**, then **Test connection**. You should see "Connected."

---

## Step 4: Load the add-on into Google Docs (10 minutes, do this once)

1. Open any Google Doc at **docs.google.com**
2. Click **Extensions > Apps Script** in the menu bar. A new tab opens with a code editor.
3. You will see a file called `Code.gs` with a few lines of starter code. Select all of it and delete it.
4. Open the `Code.gs` file from this bundle's `docs-addon` folder in any text editor. Copy ALL of it. Paste it into the Apps Script editor.
5. Now add the sidebar file: in the Apps Script editor, click the **+** next to "Files" in the left panel, choose **HTML**
6. Name it exactly `Sidebar` (capital S, no .html, the editor adds that). Press Enter.
7. Delete the starter HTML in the new file. Copy ALL of `Sidebar.html` from this bundle and paste it in.
8. Press **Ctrl+S** (Cmd+S on Mac) to save. Click the project name "Untitled project" at the top and rename it to `ScholarKit`.
9. Go back to your Google Doc tab and **reload the page** (F5)
10. A new menu called **ScholarKit** appears next to Help. Click it, then **Open ScholarKit**.
11. Google shows a scary permissions screen the first time. This is normal for personal scripts. Click through it:
    - Click **Continue** (or Review permissions)
    - Pick your Google account
    - You will see "Google hasn't verified this app." Click **Advanced** (small link, bottom left)
    - Click **Go to ScholarKit (unsafe)**. It says unsafe only because YOU wrote it and Google has not reviewed it. You are authorizing your own code.
    - Click **Allow**
12. The sidebar opens. Go to **Settings**, paste your Gemini key, click **Save**, then **Test connection**.

This add-on is now available in every Google Doc on your account via the Extensions menu? No. It is bound to this one document's script. Two options:
- Easiest: keep one "writing workspace" Doc where you draft with ScholarKit, then copy into final docs.
- Per-document: repeat steps 2 to 8 in any doc where you want it (2 minutes once you have the files).

---

## Using it

**Phrases**: pick a section from the dropdown or just search ("gap", "limitation", "hypothesis"). Click a phrase, it lands at your cursor. Replace the [bracketed] placeholders with your content.

**Rewrite**: select text in your document, click **Grab selection** (or paste text), pick a mode, click **Rewrite**. Three options appear. **Replace selection** swaps your selected text for that option. Always read AI output before keeping it, and check your school's AI policy. Capella and UK both have explicit rules on AI assistance; paraphrasing your own sentences is generally the safe zone, generating content is not.

**Word choice**: type a tired word ("shows", "big", "looked at"), get 8 academic alternatives as clickable chips.

**Similarity** (Word): add the PDFs and docs you are citing with the file picker. Click **Check document against library**. You get an overlap percentage and the exact matching passages, each with a button to highlight it in your document. Everything runs in your browser; nothing is uploaded anywhere.

**Similarity** (Google Docs): paste reference text or add other Google Docs by link. Library lasts for the current sidebar session.

Reading the score: under 5% is normal (shared terminology, citations). 5 to 15% means review the matches. Over 15% means you have unparaphrased source text in your draft.

---

## Troubleshooting

**"Could not load phrasebank.json"**: your GitHub Pages site is not live yet, or the repo is not named `scholarkit`. Test the taskpane URL from Step 2.14 directly in a browser.

**You named the repo something else**: edit `word-addin/manifest.xml` (replace every `raiderj77.github.io/scholarkit` with your actual URL, it appears 8 times) and line 1 of the script in `docs-addon/Sidebar.html` (the PHRASEBANK_URL). Then re-upload the manifest to Word.

**Gemini error "model not found"**: the model name changed. Go to aistudio.google.com, see what models are listed, and put that name (for example `gemini-3-flash`) in the Settings tab.

**Gemini error 429 "quota"**: you hit the free tier's per-minute limit. Wait a minute. For paper-writing volume you will essentially never hit the daily cap.

**Word says the add-in won't load / blank pane**: the manifest URLs must match your live Pages site exactly. Also confirm the taskpane URL opens in a normal browser.

**Highlight button says it could not locate the passage**: punctuation inside the passage differs from the normalized match. Find it manually with Ctrl+F using the first few words.

**Scanned PDFs import as empty**: they are images, not text. Run OCR on them first (Word itself can do this: File > Open the PDF, Word converts it; save as docx and import that).

**Updating the phrasebank**: edit `phrasebank.json` in your GitHub repo (pencil icon on github.com), commit, wait a minute. Both Word and Docs pick up the change on next open. Add your own phrases freely, it is your file.

---

## Legal and scope notes

- Every phrase in `phrasebank.json` was written originally for this project. It is yours to use, edit, and extend.
- Do not paste in phrase content from Ref-N-Write or the Manchester Academic Phrasebank. Both are copyrighted databases. Writing your own additions or generating them with Gemini is fine.
- This bundle is for personal use. If it ever becomes a product, the things that change are: API key moves behind a backend, the add-ins go through AppSource / Workspace Marketplace review, and the name gets a trademark check.
- The similarity checker compares against YOUR library only. It is not Turnitin and does not check the open web. It catches accidental self-overlap and unparaphrased source text from documents you give it, which is the main pre-submission risk.
