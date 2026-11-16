# Deploy Barcode Tracker to Koyeb (Free)

## What You Need
- GitHub account (free): https://github.com
- Koyeb account (free): https://www.koyeb.com

---

## Step 1: Create GitHub Repository

1. Go to https://github.com and sign up/login
2. Click "New repository"
3. Name it: `barcode-tracker`
4. Make it **Public**
5. Click "Create repository"

---

## Step 2: Install Git (if not installed)

Download from: https://git-scm.com/download/win

After installation, open **Command Prompt** and run:
```bash
git config --global user.name "Your Name"
git config --global user.email "your-email@example.com"
```

---

## Step 3: Push Code to GitHub

Open Command Prompt and run these commands one by one:

```bash
cd C:\Users\526224\.qoderwork\workspace\mvtq1h5wwdnskek1\outputs\barcode_tracker

git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/barcode-tracker.git
git push -u origin main
```

Replace `YOUR_USERNAME` with your actual GitHub username.

---

## Step 4: Deploy on Koyeb

1. Go to https://www.koyeb.com and sign up with GitHub
2. Click "Create App"
3. Select "GitHub" as deployment method
4. Choose your `barcode-tracker` repository
5. Configure:
   - **App name**: `barcode-tracker`
   - **Instance type**: Free
   - **Region**: Washington D.C. or Frankfurt (closest to your users)
   - **Build command**: `pip install -r requirements.txt`
   - **Run command**: `gunicorn app:app --bind 0.0.0.0:$PORT`
6. Click "Deploy"

---

## Step 5: Get Your Live URL

After deployment (takes 2-3 minutes), your app will be live at:
```
https://barcode-tracker-xxxxx.koyeb.app
```

Share this URL with your riders!

---

## Important Notes

**Data Storage:**
- Scan history is stored in browser localStorage (per device)
- Records are saved to DingTalk Sheets (cloud)
- No server-side database needed

**Free Tier Limits:**
- 100GB bandwidth/month
- 512MB RAM
- No sleep/timeout (always running)

**If you need to update the app:**
```bash
git add .
git commit -m "Update description"
git push
```
Koyeb will auto-deploy the changes.

---

## Troubleshooting

**Build fails:**
- Check that `requirements.txt` and `Procfile` are in the repository
- Check Koyeb logs for error messages

**App not loading:**
- Check Koyeb logs
- Make sure PORT environment variable is being used

**Riders can't access:**
- Share the full URL: `https://barcode-tracker-xxxxx.koyeb.app`
- Make sure they have internet access
