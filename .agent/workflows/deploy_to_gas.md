---
description: Deploy application to Google Apps Script
---

# Deploy to Google Apps Script

This workflow handles the deployment of the expense application to Google Apps Script using `clasp`.

## Prerequisites
- `clasp` must be installed and logged in.
- `.clasp.json` must exist in the root directory.

## Steps

1. **Push Changes**
   Upload local files to the GAS project.
   ```powershell
   clasp push
   ```
   *Note: If prompted to overwrite, confirm with 'y'.*
   // turbo

2. **Verify Deployment ID**
   Ensure you are deploying to the correct Deployment ID.
   Target ID: `AKfycbzHewMZtNIClOyN_3jWE_kPUh2wl63Ja7nqaZp_cY0UUwR7qsI3l71UXC7JTUnjosyJ`

3. **Update Deployment**
   Deploy the latest version to the Web App.
   ```powershell
   clasp deploy -i AKfycbzHewMZtNIClOyN_3jWE_kPUh2wl63Ja7nqaZp_cY0UUwR7qsI3l71UXC7JTUnjosyJ -d "Update"
   ```

## Verification
- Access the Web App URL to confirm changes are live.
- URL: `https://script.google.com/macros/s/AKfycbzHewMZtNIClOyN_3jWE_kPUh2wl63Ja7nqaZp_cY0UUwR7qsI3l71UXC7JTUnjosyJ/exec`
