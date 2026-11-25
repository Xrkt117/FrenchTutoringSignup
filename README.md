# French Tutoring Sign-Up

Sign-up page for 1-hour French tutoring sessions. Up to 3 students per session, and sign-ups and cancellations close 24 hours before a session starts.

The page is hosted on GitHub Pages. Sign-ups are saved in a Google Sheet through a small Apps Script (`Code.gs`).

## Set up the Google Sheet (once)

1. Create a new Google Sheet, then open **Extensions → Apps Script**.
2. Replace everything in `Code.gs` there with this repo's `Code.gs`, then save.
3. Pick `setup` in the function dropdown and click **Run**. Allow the permissions it asks for. This creates the **Sessions** and **Sign-ups** tabs.
4. Click **Deploy → New deployment**, choose the gear icon → **Web app**, and set:
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Copy the web app URL and paste it into `index.html` as the `API` value, then push.

Use the Sheet's own time zone (File → Settings) for session times.

## Running it

- **Add sessions:** in the **Sessions** tab, put one row per session: a start date and time (e.g. `12/2/2026 4:00 PM`), the tutor's name, and tick **Online** for online sessions. Past sessions hide themselves.
- **See who signed up:** open the **Sign-ups** tab. Delete a row there to remove a student.
- **Don't change the start time of a session that already has sign-ups.** Those sign-ups would stop counting toward it.
- **After editing `Code.gs`:** go to Deploy → Manage deployments → edit → Version: **New version**. Otherwise the old code keeps running.

## Check the rules locally

```
node test.js
```
