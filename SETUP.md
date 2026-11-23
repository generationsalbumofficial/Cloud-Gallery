# Cloud Gallery - Backend setup (Google Apps Script)

The old `Code.gs` was lost, so `backend/Code.gs` is a new backend. It is a **fresh database**: accounts and photos from the old backend will not show up here.

## 1. Create the backend (5 minutes)

1. Open https://script.google.com and click **New project**. Sign in with the Google account that should own the photos (this account's Drive stores everything).
2. Delete the default code, paste the full content of `backend/Code.gs`.
3. Left side -> **Project Settings** (gear icon) -> tick **Show "appsscript.json" manifest file in editor**. Go back to **Editor**, open `appsscript.json` and replace its content with `backend/appsscript.json`.
4. In the function dropdown choose **setup** and press **Run**. Allow the permissions (Advanced -> Go to project (unsafe) -> Allow). This creates the Google Sheet "Photo Gallery DB" and the Drive folder "Photo Gallery Files". Then run **authorizeMail** once and click Allow (needed for Forgot password).
5. **Deploy -> New deployment -> Web app**. Execute as: **Me**. Who has access: **Anyone**. Deploy and copy the **/exec** URL.
6. Open `index.html`, find `API_URL` and paste the URL.
7. Upload the website files to GitHub Pages as before.

Changed the code later? Use **Deploy -> Manage deployments -> Edit (pencil) -> Version: New version -> Deploy**. The URL stays the same.

## Login system (same as Generation Album)

- **Sign up:** name, email, password. The account is created and logged in at once. There is no OTP and no email verification.
- **Log in:** 5 wrong passwords lock that email for 15 minutes. A login stays active for 30 days; Logout ends it on the server.
- **Forgot password:** a new temporary password is e-mailed (valid 60 minutes, max 3 mails per hour). The old password keeps working until the new one is used. Change it later from Menu > Account.
- **Admin:** the account that signs up with `ADMIN_EMAIL` gets the admin panel. Create it yourself right after deploying.

Updating an already deployed version: paste the new `Code.gs`, run **authorizeMail**, then Deploy > Manage deployments > Edit > New version > Deploy. Existing accounts and passwords keep working; everyone just has to log in once again.

## 2. Settings you may want to change

`backend/Code.gs` (top of the file): `ADMIN_EMAIL`, `DEFAULT_LIMIT_GB`, `MAX_ITEMS_PER_USER` (1000, hidden from users), `MAX_VIDEO_BYTES` (100 MB).

`index.html` -> `CFG`:

| Setting | Meaning |
|---|---|
| `VIDEO_MAX_MB` | Max size of a video after compression (default 100) |
| `VIDEO_PICK_MAX_MB` | Biggest original file the browser will try to compress (default 600) |
| `VIDEO_BPP` | Quality target. Higher = better quality, bigger file (default 0.14, visually lossless) |
| `ADMIN_EMAIL` | The contact e-mail shown in the app (still a placeholder) |

## 3. How video upload works

- **Compression** runs in the user's browser (WebCodecs, library: Mediabunny, loaded from a CDN). The resolution and frame rate are kept. Audio is kept as it is whenever possible.
- A video that is already efficient (or where compression would save less than 8%) is uploaded **untouched**.
- Only when a video would not fit into 100 MB at full quality, the bitrate is lowered (and the resolution slightly, if really needed) so that it fits.
- With several files selected, the next file is compressed while the current one uploads.
- The upload is sent in 4 MB chunks (an Apps Script request cannot carry 100 MB at once). A failed chunk is retried automatically.
- Playing uses Google Drive's player inside the viewer. Right after an upload, Google needs a short time to create the thumbnail and to process the video.

## 4. Good to know

- Compression needs a browser with WebCodecs (Chrome, Edge, Safari 16.4+, recent Firefox). If it is not available, videos up to 100 MB are uploaded as they are.
- Forgot-password e-mails are sent with `MailApp` (about 100 e-mails per day on a free Google account).
- Apps Script has daily quotas (for example URL Fetch calls). Every 4 MB chunk uses one call.
