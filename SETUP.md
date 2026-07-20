# Department of Defence External Recruitment Portal

This package connects the homepage, the 313-position vacancy register, the online application form, the success page, and a Google Apps Script backend.

## Website files

Upload these files to the root of the GitHub repository:

- `index.html`
- `vacancies_by_divisions.html`
- `application-form.html`
- `success.html`
- `vacancies-data.js`
- your existing `logo.png`
- your existing `background.jpg`

The homepage and every vacancy **Apply** button now open `application-form.html`. Position details are passed in the URL and filled automatically. The application form also loads the same 313-position register from `vacancies-data.js`.

## Create the Google Apps Script backend

1. Go to <https://script.google.com> and create a **New project**.
2. Rename the project to **DOD External Recruitment Backend**.
3. Replace the default `Code.gs` content with the complete supplied `Code.gs` file.
4. Open **Project Settings**, enable **Show appsscript.json manifest file in editor**, and replace its contents with the supplied `appsscript.json`.
5. Return to `Code.gs` and run this once from the editor:

   ```javascript
   setupRecruitmentBackend('your-recruitment-mailbox@example.gov.pg')
   ```

   Replace the example address with the official recruitment mailbox. Google will ask you to authorize Drive, Sheets, and email access. The setup function creates a private applications spreadsheet and a private uploads folder in the account running the script.

6. Click **Deploy → New deployment → Web app**.
7. Set **Execute as** to **Me**.
8. Set **Who has access** to **Anyone**.
9. Click **Deploy**, authorize it, and copy the URL ending in `/exec`.
10. In `application-form.html`, find this line near the top:

   ```html
   <meta name="dod-google-script-url" content="YOUR_URL_HERE" />
   ```

   Replace only the `content` value with the new `/exec` URL.

## Test before publishing

1. Upload the website files to GitHub and wait for GitHub Pages to deploy.
2. Open `vacancies_by_divisions.html` from the public website.
3. Select a vacancy and click **Apply**.
4. Confirm that division, position title, reference number, position code, cost account, and class are filled correctly.
5. Submit one test application using small test PDF and image files.
6. Confirm all of the following:

   - the success page shows the application reference;
   - a row appears in the **Applications** spreadsheet;
   - an applicant folder appears inside **DOD External Recruitment Uploads**;
   - the applicant receives an acknowledgement email;
   - the recruitment mailbox receives the internal notification.

## Important operational notes

- Applicant uploads remain private in the Google Drive account that owns the Apps Script deployment.
- Do not publicly share the applications spreadsheet or upload folder.
- The public `/exec` address is required for external applicants and is not a password or secret.
- When `Code.gs` changes, use **Deploy → Manage deployments → Edit → New version → Deploy**. The `/exec` URL can remain the same.
- Google Apps Script and Mail quotas apply. Before a large recruitment launch, confirm the daily email and storage limits of the account used for deployment.
