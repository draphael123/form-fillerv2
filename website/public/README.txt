Place docufill-extension.zip here for the website download button.

To create the zip, from the docusign-autofill/extension/ folder run:
  zip -r ../website/public/docufill-extension.zip . -x "*.DS_Store" -x "generate-icons.js"

On Windows (PowerShell), you can use Compress-Archive or a similar tool to zip the extension folder contents.
