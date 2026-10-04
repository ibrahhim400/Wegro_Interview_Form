# Interview Scheduler (Google Apps Script)

Ekta internal form (Google Form lagbe na) jeta:
- Google Sheet-er existing tab-e column-wise data save kore
- Candidate + interviewer-der kache standard Google Calendar invitation pathay (Google Meet "Open" access shoho)
- CV attach kore
- Chaile WhatsApp group / personal number-e (Green-API diye) invitation details pathay

> **Eta ekta Google Apps Script project.** Code GitHub-e thakle-o, chalanor jonno
> Apps Script-e (script.google.com) deploy kora lagbe — GitHub shudhu code store/backup kore.

## Files

- `Code.gs` — Apps Script backend (Sheet-er menu-er jonno form-o ekhanei ache)
- `index.html` — standalone page, GitHub Pages-e host korar jonno (same form, `fetch()` diye `Code.gs`-er Web App URL call kore)

## GitHub Pages-e host korte

1. `Code.gs` deploy kora thakte hobe (Web app, "Anyone" access), tar `/exec` URL ta lagbe.
2. `index.html`-er bhitor ei line ta khuje nijer URL diye replace koro:
   ```javascript
   var APPS_SCRIPT_URL = 'https://script.google.com/macros/s/XXXXXXXX/exec';
   ```
3. GitHub repo-te upload koro, **Settings → Pages → Branch: main → Save**.
4. Kichu khon por `https://<username>.github.io/<repo-name>/` link-e form ta live thakbe.

> `Code.gs`-e notun code push/paste korle, tar সাথে `index.html`-er `APPS_SCRIPT_URL` same thakbe jotokkhon na deployment URL change hoy (New version deploy korle URL change hoy na, tai eta ekbar-i set korle hobe).

## Setup (notun kore install korte)

1. Tomar Google Sheet-e **Extensions → Apps Script** e jan.
2. Default code muche ei repo-r `Code.gs` er code paste koren, save koren.
3. **Script Properties set koren** (Project Settings ⚙️ → Script Properties → Add property):

| Key | Value (example) |
|---|---|
| `SPREADSHEET_ID` | tomar Google Sheet-er ID |
| `COMPANY_NAME` | `WeGro` |
| `DEFAULT_INTERVIEWERS` | `["a@company.com","b@company.com"]` |
| `GREEN_API_ID_INSTANCE` | Green-API-r idInstance |
| `GREEN_API_TOKEN` | Green-API-r apiTokenInstance |
| `WHATSAPP_GROUPS` | `[{"id":"hr","name":"HR Internal Team","groupId":"xxxxxxxxxxxxxxxxxx@g.us"}]` |

   WhatsApp lagbe na hole `GREEN_API_*` ar `WHATSAPP_GROUPS` faka rakhle-o cholbe — shudhu WhatsApp message jabe na, baki shob thik thakbe.

4. Google Calendar API enable koren (Apps Script → Services → + → Google Calendar API).
5. `authorize` function ekbar **Run** kore shob permission Allow koren.
6. **Deploy → New deployment → Web app** → Execute as: **Me** → access thik kore Deploy koren.
7. Sheet-er `Interviews` tab-e ei header gulo rakhun:
   `Employee Name | Email | Mobile | Position | Interviewers | Interview Date | Time | Interview Type | Status | Area | Notes | CV Link | Invite Status`

## Security note

Ei repo-te kono API key/token/Sheet ID **nai** — shob `PropertiesService.getScriptProperties()` theke ashe, jeta Google-er nijer encrypted storage, GitHub-e kokhono ashbe na. Repo Public thakleo tai kono credential leak hoy na.
