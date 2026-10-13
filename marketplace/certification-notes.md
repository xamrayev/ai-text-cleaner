# Notes for certification

Paste the text below into Partner Center → Offer → Review and publish → **Notes for certification**.

---

AI Text Cleaner is a Word task pane add-in with three ribbon commands. It needs no account, sign-in, license key, payment or external service. All processing is local JavaScript using the Word API, and no data is sent to any server.

SUPPORTED PLATFORMS: Word on the web, Word on Windows and Word on Mac (Microsoft 365). Highlighting requires WordApi 1.3, and restoring the user's own highlighting exactly requires WordApi 1.4. On older Word versions the "AI traces" button is disabled and explained in a tooltip, while cleaning still works.

TEST DOCUMENT: paste the following line into a blank Word document. It contains a zero-width space between "Sa" and "lom", an em dash, curly quotes, an ellipsis and Markdown symbols:

```
## Heading
Sa​lom world — “quoted” text… and **bold** words
```

(If the zero-width character is lost while copying, the dash, quotes, ellipsis and asterisks are enough to test.)

TEST STEPS

1. Open the Home tab. A group named "AI Cleaner" contains four buttons: Cleaner, AI traces, Quick clean, Remove colors.
2. Click **AI traces**, with nothing selected. Expected: words with problems are highlighted in yellow, and a small dialog reports how many spots were found. The dialog closes with OK or by itself after about 6 seconds.
3. Click **Quick clean**. Expected: the text becomes `Heading` / `Salom world - "quoted" text... and bold words`, cleaned words are highlighted in green, and a dialog reports the number of changes.
4. Click **Remove colors**. Expected: all yellow and green highlighting is removed, and a dialog reports how many spots were restored.
5. Press Ctrl+Z (⌘Z on Mac) several times to restore the original text.
6. Click **Cleaner** to open the task pane.
   a. On the "Document" tab, click "AI traces". A list of found spots appears, and clicking an item selects that word in the document.
   b. Click "Clean". A results table shows counts per rule.
   c. Open the "Text" tab, paste the test line into "Input text" and click "Clean". The cleaned text appears and can be copied or inserted into the document.
   d. Expand "Cleaning rules" and switch any rule off. It is not applied on the next run. "Restore defaults" resets the rules.
   e. Use the EN / RU / UZ selector in the top-right corner to switch the interface language. By default the add-in follows the Office display language.
7. The footer links open the Privacy, Terms and Support pages.

KNOWN LIMITATIONS (by design, also documented on the support page): Word search does not always reach characters inside fields, text boxes or comments. When this happens, the panel reports the number of remaining characters and suggests Text mode.

Contact for reviewers: {{SUPPORT_EMAIL}}
