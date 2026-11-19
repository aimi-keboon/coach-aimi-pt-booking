# Coach Aimi — Personal Training Booking

Static GitHub Pages landing site + Google Apps Script booking app + Google Sheets control panel + Google Calendars. No hosting/database platform beyond these Google services and GitHub Pages.

## Included

- **`index.html`**: static branded GitHub Pages landing page with a Book button.
- **`apps-script/Code.gs`**: backend that checks all enabled availability calendars but creates each booking only on the designated calendar, plus Google Sheet configuration, email sign-in, booking/rescheduling/cancelling, prepaid credits, and mail notifications.
- **`apps-script/Booking.html`**: the interactive client booking app, hosted by Google Apps Script HTML Service.
- **`apps-script/appsscript.json`**: set to `Asia/Kuala_Lumpur` time zone.

## Required Google accounts and permissions

Plan: use **`keboon@keboon.net`** to create/deploy the Apps Script and (optionally) own the Sheet and daily trigger, **if it is a working Google/Workspace account**. The deploying account **must have permission to create/edit/delete events on the booking destination calendar (`aimi@keboon.net`)**. For all other calendars, it needs sufficient **read access** to check busy events; write access to those calendars is not required. Sharing the destination calendar with "See all event details" is insufficient—grant **Make changes to events**, or stronger.

Booking notifications will go to both the registered client's email and **`aimi@keboon.net`**. `MailApp` sends using the **account executing the script**; setting a `Reply-To` address **does not** change the sender. To send from `keboon@keboon.net`, deploy as that account. You do **not** need any cron job for immediate booking emails: they are sent during the booking/reschedule/cancel operation. A daily audit cron is optional.

## 1. Deploy the Apps Script

1. Log in to `https://script.google.com` as the account that can edit all target calendars (ideally `keboon@keboon.net`), and create a **new standalone Apps Script** project.
2. Replace `Code.gs` contents with `apps-script/Code.gs`.
3. Add a new **HTML** file named `Booking`, and paste in `apps-script/Booking.html` (do not add a second `.html` extension to its Apps Script file name).
4. In Project Settings, enable **Show appsscript.json manifest file in editor**. Replace `appsscript.json` with the contents of `apps-script/appsscript.json`.
5. Save. Select `setupProject_` from the function menu and click **Run**. Grant the required permissions. It creates a private spreadsheet (or migrates the existing Clients tab to include CreditsGranted). Open **Executions / Execution log** to get its URL. Running it again does not overwrite your data.
6. In the **Settings** tab, confirm `BookingCalendarId = aimi@keboon.net` (or the exact ID of that calendar if different). In the **Calendars** tab, enter the IDs of your *other* calendars and set **Enabled** to `TRUE` for each you want to block busy slots. The booking destination is checked automatically, even if not listed in the Calendars tab. Add as many additional availability calendars as needed.
7. Select `checkSetup_` in Apps Script, **Run**, and verify it reports all availability calendar IDs, the destination ID, and the sheet URL. Make a test booking after publishing to confirm the deployment owner has write access to **only the destination as required**.
8. Choose **Deploy → New deployment → Web app**. Use **Execute as: Me** and **Who has access: Anyone**. Confirm that your Google/Workspace policy permits anonymous access. Copy the deployed `/exec` URL. This grants visitors access to the _limited booking functions_ as the deployer; **never** add publicly callable admin functions that expose data.
9. If you change the Apps Script source later, **Deploy → Manage deployments → Edit → New version**. Saving code alone does not update the production `/exec` deployment.

## 2. Configure the Google Sheet

You can do all everyday admin work here, without a new admin dashboard:

### Clients

| Name | Email | Phone | Active | CreditsGranted |
| --- | --- | --- | --- | --- |
| Example Client | example@example.com | `0123456789` | TRUE | 10 |

- **Replace the example with a real client; do not enter sample data publicly.**
- The `Phone` column is formatted as text; maintain one unique phone per client. You can type `0123456789`, `60123456789`, or `+60123456789`; login normalizes these Malaysian formats.
- **CreditsGranted** is the *cumulative total credits you have allocated* to a client. Set to `10` to allow up to 10 one-hour sessions, and change to `20` if you sell 10 more credits later (do not set it back to `10`). The app calculates the current balance automatically by subtracting all **CONFIRMED** bookings, including completed/no-show sessions. It never edits the credit grant in your sheet.
- A new confirmed booking uses one credit; moving the booking costs no extra credit; cancelling **at least 24 hours before** the start returns the credit automatically, because the booking becomes CANCELLED. In the last 24 hours, reschedule and cancel controls are blocked on both frontend and server. A no-show uses up the credit without any special process. To correct a balance manually, change CreditsGranted. Keep the Bookings history for correct accounting; never delete rows to make the balance look different.
- New clients cannot self-register. You enroll them manually; they then sign in via phone plus an email code sent to the saved address. Set `Active=FALSE` to block account access without deleting bookings.
- Make the spreadsheet private. Never commit actual client names, phones, emails or OTP secrets to GitHub.

### Schedule

Default `TRUE` working days: **Monday, Tuesday, Thursday, Friday, Saturday**. Default off days: **Wednesday and Sunday** (`FALSE`).

`FirstHour=7`, `LastStartHour=22` means possible starts are **07:00, 08:00, ... 22:00**, 1 hour long; last session is **22:00–23:00**.

To take an additional off day, change that row's **Open** value to `FALSE`; to work on an off day, change it to `TRUE`. You can also set a different first/last hour per day (whole hours, 0–22).

**Previously confirmed bookings are never automatically deleted** when you change the schedule. Clients still see them in My bookings. New bookings and reschedules follow the latest schedule.

To take just one date off, create a **Busy all-day event on any calendar being checked for availability** for that date. Existing bookings remain valid; future slots on that date become unavailable. Mark the event *Busy*, not *Free*.

### Calendars

**This is availability-only.** Every enabled row blocks busy slots; the app does **not** create booking events on these additional calendars. The destination `aimi@keboon.net` is also checked for conflicts automatically.

| CalendarId | Label | Enabled |
| --- | --- | --- |
| `your-other-calendar-id` | Personal / other work | TRUE |
| `another-calendar-id` | Optional third calendar | FALSE |

Get a Calendar ID in Google Calendar on desktop: **Settings → [calendar name] → Integrate calendar → Calendar ID**. The account running Apps Script must be able to read each enabled availability calendar (including its busy events). To check an additional calendar, add a row or set Enabled to `TRUE`. To stop checking it, set `FALSE`. **This does not alter existing confirmed bookings.**

**Booking destination (not an availability-only calendar):** Set `Settings → BookingCalendarId` to `aimi@keboon.net`. Every new booking creates **exactly one event** there, and only that event is moved/deleted when a client reschedules/cancels. Apps Script must have **write access** to this destination. If the calendar is a secondary calendar with a non-email ID, use the exact ID instead of the email address.

**All enabled busy-check calendars + the destination** are used when deciding which slots are available. If you do not add your second calendar to the Calendars tab, the app cannot detect busy times on it.

### Settings

- `CoachEmail`: `aimi@keboon.net` (receives booking, rescheduling and cancellation notifications)
- `BookingCalendarId`: `aimi@keboon.net` (the **only** calendar that receives new booking events)
- `ReplyToEmail`: `aimi@keboon.net`
- `BookingHorizonDays`: `60`
- `RescheduleCutoffHours`: `24`

## 3. Connect GitHub Pages

1. Create a **public GitHub repository** called `coach-aimi-pt-booking` under your GitHub account (`aimi-keboon`), or choose another name you own. Upload the files from this package to its default branch.
2. In root **`index.html`**, replace `PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE` with the deployed Apps Script `/exec` URL. The GitHub Pages **Book your session** button will open the real booking experience.
3. In GitHub repository **Settings → Pages**, select **Deploy from a branch**, then `main` / `/(root)`, Save.
4. After Pages publishes, your landing page will generally be `https://aimi-keboon.github.io/coach-aimi-pt-booking/` (if you keep that repository name).

Why a separate booking page? `google.script.run` works securely *inside* Apps Script HTML Service. GitHub Pages cannot directly call it, and direct CORS fetch calls to Apps Script can be unreliable. Apps Script iframe embedding also has reported recent breakages even with `ALLOWALL`. A simple branded link avoids those limitations.

## 4. Test these cases before sharing the link

1. Add **one real test client** with an accessible email and phone and set `CreditsGranted=2`.
2. Sign in with phone → request OTP → verify email code.
3. Try a Wednesday and Sunday: no available slots; all confirmed sessions already booked for that day remain visible in My bookings.
4. Book Monday **7–8 AM** (on a future date) and verify **one event only**, in the calendar identified by `BookingCalendarId` (`aimi@keboon.net`), with **no booking event** created on other calendars.
5. Confirm both the client and `aimi@keboon.net` receive notifications. Check spam and sender identity.
6. Verify the time disappears from availability; another client should not be able to double-book it.
7. Reschedule a session more than 24 hours away. Verify the old event disappears and the new event appears **only in `aimi@keboon.net`**, and both parties receive emails.
8. Cancel a booking more than 24 hours out, verify its credit returns. Book again and check the displayed balance. Try to edit/cancel within 24 hours; server must reject the request and the credit remains spent. Booking with zero remaining credits must be rejected even when a slot is free.
9. Change a day to `Open=FALSE`. Already-confirmed sessions remain booked and visible; future open slots on that weekday disappear.
10. Make a manual Busy event on any enabled availability calendar: the overlapping slot must disappear from availability even though bookings are created only on the destination.
11. Test after you change availability calendar IDs: already-confirmed bookings are preserved. If changing `BookingCalendarId`, verify new bookings go to the new destination and old booking references still point to their original calendar.

## Optional cron / daily audit

Run **`installAuditTrigger_`** from the Apps Script editor while signed into the desired automation account (`keboon@keboon.net`). It creates a daily ~8 AM MYT trigger that checks upcoming bookings for missing events on the booking destination and emails the coach **only if it finds problems**. Triggers execute under the account that creates them. This is not necessary for immediate notifications, and it does not continuously mirror manual edits made directly inside a Google Calendar.

## Security and operational limitations

- **Phone-only login is not secure.** The email OTP is required before viewing any private bookings or booking/editing sessions. There is no SMS OTP provider in this Google-only setup.
- The public GitHub repository contains **source code only**. Data, authentication codes, spreadsheet IDs and private calendar IDs remain in your own Google account / Apps Script properties / private Sheet. Do not publish real client information.
- The deploying account needs **edit access to the booking destination** and read access to other enabled calendars. Apps Script uses `LockService` and rechecks all enabled calendars plus confirmed booking rows before saving to reduce double bookings. External edits made during booking and partial Google service failures are not a fully transactional database; if synchronization errors occur, check the Apps Script logs, calendar events and Bookings sheet.
- Credits are counted from booking rows rather than decremented from the sheet, which avoids double-spending and makes refunds automatic. A cancelled booking does not consume a credit; confirmed sessions from the past continue consuming credits even if the client does not show up. Do not remove the history rows.
- The 24-hour rule is checked **server-side**, against the original start time for a reschedule/cancel, and also against the new selected time for reschedules.
- Modifying an event manually in Calendar does **not** automatically edit the corresponding Sheet booking. Manage client reschedules from the booking website. The optional audit warns of missing destination events but does not repair changes.
- MailApp sends from the web app's executing account; `ReplyToEmail` only controls replies. Sending email is subject to Google's daily sending quotas and anti-abuse limits. Booking confirmations are written before notification emails are attempted, so if email delivery fails a booking may still be valid. Check Apps Script execution logs when investigating delivery.
- The web app requires public anonymous access for enrolled clients without Google accounts. Check Google Workspace/domain policy, OAuth consent, account quotas and your privacy obligations before using real client data.

## Source references

- Apps Script web apps: https://developers.google.com/apps-script/guides/web
- Google Calendar Apps Script API: https://developers.google.com/apps-script/reference/calendar/calendar-app
- LockService: https://developers.google.com/apps-script/reference/lock
- MailApp: https://developers.google.com/apps-script/reference/mail/mail-app
- Apps Script quotas: https://developers.google.com/apps-script/guides/services/quotas