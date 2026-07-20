# Department of Defence PNG External Recruitment Portal

Public-facing recruitment website for Department of Defence position vacancies and online applications.

## Website

The GitHub Pages site is published from the `main` branch and repository root:

<https://depaulkuyama-tech.github.io/dod-recruitment-portal/>

## Main files

- `index.html` — homepage
- `vacancies_by_divisions.html` — searchable position register
- `vacancies-data.js` — shared vacancy data
- `application-form.html` — external application form
- `success.html` — submission confirmation page
- `Code.gs` and `appsscript.json` — Google Apps Script application receiver
- `SETUP.md` — backend deployment and testing instructions

## Deployment

GitHub Pages must use **Deploy from a branch**, with branch `main` and folder `/ (root)`.

Follow [SETUP.md](SETUP.md) before accepting live applications. The Google Apps Script deployment, applications spreadsheet, and upload folder must remain under an authorized Department of Defence account.
