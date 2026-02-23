Place docufill-extension.zip here for the website download button.
Optional: Add og.png (e.g. 1200x630px) for social sharing previews (Twitter, LinkedIn, etc.).

To create the zip, from the docusign-autofill/extension/ folder run:
  zip -r ../website/public/docufill-extension.zip . -x "*.DS_Store" -x "generate-icons.js"

On Windows (PowerShell), you can use Compress-Archive or a similar tool to zip the extension folder contents.
