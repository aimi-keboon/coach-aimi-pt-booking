# Coach Aimi PT Booking — GitHub Pages + Google Apps Script

**The full customer-facing HTML, CSS and JavaScript live in `index.html` on GitHub Pages.** The Apps Script web app is only a backend: `Code.gs` serves a tiny invisible messaging bridge (`google.script.run` + `postMessage`). No customer interface is hosted by Apps Script.

## Features

- One-hour personal training: 7:00 AM through 10:00 PM starts (last session 10–11 PM), Malaysian timezone.
- Wednesday/Sunday off by default; change off days and hours in the private **Schedule** sheet. Changing hours or closing a day never cancels already confirmed bookings.
- Check **all enabled calendars** plus destination calendar for conflicts; **create booking events on `aimi@keboon.net` only**.
- Private Google Sheet **Clients** tab with `Name`, `Email`, `Phone`, `Active`, `CreditsGranted`. One credit spent per confirmed booking; a cancellation more than 24 hours ahead returns the credit. No-show/late sessions retain their credit use.
- Registered client enters phone, receives email OTP, sees credits/bookings/availability, books, cancels or reschedules at least 24 hours before a session.
- Email notifications to `aimi@keboon.net` and client for bookings/cancellations/reschedules; optional daily audit trigger.

## Project files

- `index.html`: simple public GitHub Pages homepage with links for Client Login and Coach Login.
- `admin.html`: version-controlled source UI for the coach dashboard. The deployed live admin dashboard is now served by Apps Script at `/exec?page=admin`.
- `apps-script/Code.gs`: Google Calendar/Sheets/Email backend plus bundled coach and client UI. **Paste into Google Apps Script's `Code.gs`.**
- **No JSON file to copy:** the backend already has `TZ: 'Asia/Kuala_Lumpur'` in `Code.gs`. The Apps Script editor manages its own hidden manifest automatically.

**You do not create `Booking.html` or `appsscript.json` inside Apps Script.** Only paste `Code.gs`. Google Apps Script maintains its own manifest automatically (hidden by default).

## Setup

1. Log into [Google Apps Script](https://script.google.com) as the Google account that should execute bookings and email (you suggested `keboon@keboon.net`). Create a standalone project and paste `apps-script/Code.gs` into `Code.gs`.
2. Open **Project Settings (gear icon) → Time zone** and choose **GMT+08:00 Kuala Lumpur / Asia/Kuala_Lumpur** (or a matching GMT+08:00 time zone). You do not need to show or edit `appsscript.json`. Make sure **Enable Chrome V8 runtime** is on.
3. Save `Code.gs`; select **`setupProject`** (without `_`) from the Apps Script Run dropdown, click **Run**, authorize, and copy the Google Sheet URL from **Execution log**. Do not share this spreadsheet publicly.
4. Add each client in the `Clients` tab. Enter phone numbers as **plain text** (e.g. `0123456789`) and positive whole-number `CreditsGranted` values. Do not delete past rows from `Bookings` because the credit balance depends on them.
5. In `Settings`, leave `BookingCalendarId` set to `aimi@keboon.net` **only if that is the actual Google Calendar ID**. Confirm using **Google Calendar → Settings → [calendar] → Integrate calendar → Calendar ID**. In `Calendars`, enter all other calendars to check for conflicts with `Enabled` = TRUE. They need read access; destination needs permission to modify events. Do not enable the Malaysia holidays calendar unless you want holidays to block sessions.
6. Run `checkSetup` in Apps Script and confirm you can access all calendars.
7. **Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone.** Save/copy the final `/exec` URL. Confirm your Workspace policy allows anonymous web app use. Anyone only gains access to the *public* actions; booking details remain behind registered-phone email OTP verification.
8. The GitHub `index.html` already links to your `/exec?page=client` and `/exec?page=admin` pages. No URL edit is necessary unless the web app deployment URL changes.
9. On GitHub, go to repository **Settings → Pages → Build and deployment → Deploy from branch → main / (root)**. Open the resulting GitHub Pages site and test sign-in, booking, rescheduling, cancellation, calendar conflicts, credits, and both email notifications using a test client.
10. Optionally run `installAuditTrigger_` from the Apps Script editor. Audit is supplemental; normal booking emails send when each action happens. Gmail/Apps Script email quotas still apply.

## How login pages work (reliable top-level hosting)

- GitHub Pages remains the **public homepage**, with simple Client Login and Coach Login links.
- Both links open the existing Google Apps Script web app as a **top-level page** with `?page=client` or `?page=admin`.
- All booking and admin forms still use the same HTML/CSS/JavaScript design, now served via `HtmlService` from `Code.gs`. The backend uses Google's supported `google.script.run` inside the top-level Apps Script web app.
- No embedded iframe, no cross-origin fetch, no separate HTML files to create in the Apps Script editor.
- After each `Code.gs` change, paste the new code into Apps Script, save, and **Deploy → Manage deployments → Edit → New version → Deploy**.
- The fixed public homepage points at the existing `/exec` URL; if a deployment gets a completely new URL, update its two login links.

## Access and security

- Never store Google Calendar credentials, client phone/email lists or client booking data in GitHub; GitHub Pages is public.
- The Apps Script runtime reads private Google Sheets and calendars as the deployment owner. The `Code.gs` embedded bridge accepts calls only from the configured GitHub Pages origin and only to explicitly whitelisted methods.
- A phone number alone does **not** reveal bookings; the app emails a 6-digit one-time code to the email already recorded in the private Clients sheet. The code expires after 5 minutes; subsequent actions use a short-lived session token.
- Apps Script is not intended for very high traffic. Rate limits, email quotas, and browser compatibility can affect operation; perform acceptance tests before real client rollout.
- The booking events live only on the destination calendar; other selected calendars only block unavailable slots.

## Troubleshooting

- **App login opens error or blank page:** ensure your deployed Apps Script version contains the latest bundled `Code.gs` and visit `/exec?page=admin` or `/exec?page=client` directly.
- **Login page shows an older page:** redeploy Apps Script as a **New version** under your existing web app deployment. It no longer uses an iframe.
- **Calendar not accessible:** share the correct calendar with the executing Google account. It must be an editable owner/writer of the destination calendar and able to read each blocker.
- **Email missing:** check MailApp authorization, quotas, spam folders, client email entry and whether the sender account is correctly configured.
- **Date/time shifts:** set the Apps Script project timezone in Project Settings to GMT+08:00, and set the spreadsheet timezone to Malaysia (GMT+08:00).

## Coach Admin dashboard

- Open **https://aimi-keboon.github.io/coach-aimi-pt-booking/** and choose Coach Login (direct Apps Script `?page=admin`).
- Request the coach-only one-time sign-in code. It is sent to `CoachEmail` in the private Settings sheet (initially `aimi@keboon.net`).
- Add a client with name, email, phone number, Active status, and starting credits. Updating the same phone number keeps existing credits.
- For existing clients use **+ Add credits**. This adds to granted credits; spent credits remain counted in booking history.
- Add extra availability calendars using their Calendar ID. An **email address can serve as the ID of a Google account's primary calendar** if the Apps Script executor already has access. For a secondary calendar copy its full Calendar ID from Google Calendar Settings → Integrate calendar.
- In Google Calendar, **share the other calendar** with the executing Apps Script account (`keboon@keboon.net`, if that is the account you deployed under) with enough permissions to see busy events. Entering an email address does not itself share or authorize calendar access.
- Enable or disable extra calendars from the admin page; **the booking destination (`aimi@keboon.net`) is always checked and remains the only calendar receiving new booking events**.
- Do not share the private Sheets URL with clients. Public GitHub Pages files contain no passwords or admin client records.

**After changing Code.gs:** Save it in the Apps Script editor, then go to **Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy**. The /exec endpoint continues using the old code until you redeploy. Keep the existing /exec URL.

**Security note:** This uses email OTP and a six-hour server-side admin session; never put a static admin password in a public GitHub Pages file. Treat the coach email inbox as sensitive because possession of a code allows administrative access. Test direct Apps Script login and booking workflows before going live.
