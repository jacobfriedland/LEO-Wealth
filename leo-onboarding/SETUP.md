# Setup Instructions

## Step 1 — Connect the submission email (Formspree)

The form sends submissions to your email via **Formspree** (free, no server needed).

1. Go to [formspree.io](https://formspree.io) and create a free account
2. Click **+ New Form**, name it "Leo IRP", and set the email to `jacob.friedland@leowealth.com`
3. Copy the form endpoint — it looks like: `https://formspree.io/f/abcd1234`
4. Open `index.html` and find this line near the top of the `<form>` tag:
   ```
   action="https://formspree.io/f/REPLACE_WITH_YOUR_FORM_ID"
   ```
5. Replace `REPLACE_WITH_YOUR_FORM_ID` with your actual form ID (e.g. `abcd1234`)

When a client submits the form, you'll receive an email at `jacob.friedland@leowealth.com` with all their answers, score, risk profile, and signature.

To change the recipient email later, log into Formspree and update it there — no code change needed.

---

## Step 2 — Publish to GitHub Pages (get a live link)

1. Push this folder to GitHub (it's already in the repo)
2. On GitHub, go to the repo → **Settings** → **Pages**
3. Under "Source", select **Deploy from a branch**, pick `master`, folder `/leo-onboarding`
4. Save — GitHub will give you a URL like:
   ```
   https://jacobfriedland.github.io/leo-onboarding/
   ```
   This is the link you email to clients.

---

## Step 3 (later) — Switch to a custom domain

When you're ready to use something like `forms.leowealth.com`:

1. In the repo, create a file `leo-onboarding/CNAME` containing just:
   ```
   forms.leowealth.com
   ```
2. Ask your IT/DNS team to add a CNAME record: `forms.leowealth.com` → `jacobfriedland.github.io`
3. In GitHub Pages settings, enable "Enforce HTTPS"

That's it — the same form, now at a professional URL.

---

## Changing the recipient email later

Log into formspree.io → your form → Settings → change the email. No code change needed.
