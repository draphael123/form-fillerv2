# DocuFill — DocuSign Autofill

Chrome Extension (Manifest V3) that autofills DocuSign documents from a CSV/Excel source file, plus a static landing page.

## Project structure

```
docusign-autofill/
├── extension/          # Chrome extension
│   ├── manifest.json
│   ├── popup.html, popup.js
│   ├── content.js, background.js
│   ├── assets/         # icon16.png, icon48.png, icon128.png
│   └── generate-icons.js  # optional: node generate-icons.js (requires canvas)
└── website/            # Landing page
    ├── index.html
    └── public/         # Put docufill-extension.zip here for download link
```

## Load in Chrome

1. Open `chrome://extensions`
2. Turn **Developer mode** ON (top right)
3. Click **Load unpacked**
4. Select the `extension/` folder
5. Pin the DocuFill icon from the toolbar

## Package for download

From the repo root (`docusign-autofill/`), run:

- **Windows (PowerShell):** `.\package-extension.ps1` — creates `website/public/docufill-extension.zip` (excludes `generate-icons.js`).
- **Mac/Linux:** From `extension/`: `zip -r ../website/public/docufill-extension.zip . -x "*.DS_Store" -x "generate-icons.js"`

## First-use workflow

1. Upload your CSV/Excel in the **Data** tab.
2. Open a DocuSign document (e.g. provider agreement).
3. **Mapping** tab → enter a template name → **Start Capturing Fields** → click each DocuSign field → map to CSV column → **Save Template**.
4. **Fill** tab → select template → choose record → **⚡ Autofill Document**.
