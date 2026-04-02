# Deployment Guide — Pricing Simulation

## Option A: GitHub + Vercel (Recommended — easiest)

### Step 1: Create a GitHub repo

1. Go to **https://github.com/new**
2. Name it `pricing-simulation`
3. Keep it **Public** (or Private — both work with Vercel)
4. Do **NOT** initialize with README (we already have one)
5. Click **Create repository**

### Step 2: Push from your machine

Open a terminal, navigate to the downloaded `pricing-simulation` folder, and run:

```bash
cd pricing-simulation

git init
git add .
git commit -m "Initial commit: Pricing simulation app"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/pricing-simulation.git
git push -u origin main
```

Replace `YOUR_USERNAME` with your GitHub username.

### Step 3: Deploy on Vercel

1. Go to **https://vercel.com** and sign in with GitHub
2. Click **"Add New…" → Project**
3. Find and **Import** `pricing-simulation` from your GitHub repos
4. Vercel auto-detects Vite — leave all settings as default
5. Click **Deploy**
6. In ~30 seconds you'll get a live URL like `https://pricing-simulation-xxx.vercel.app`

That's it! Every future `git push` will auto-deploy.

---

## Option B: Vercel CLI (no GitHub needed)

```bash
# Install Vercel CLI
npm i -g vercel

# Navigate to project folder
cd pricing-simulation

# Deploy (follow prompts to log in)
vercel

# For production deployment:
vercel --prod
```

---

## After Deployment: Set Up Google Sheets

1. Create a new **Google Sheet**
2. Go to **Extensions → Apps Script**
3. Delete default code, paste contents of `google-apps-script.js`
4. Click **Deploy → New deployment**
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Click **Deploy**, authorize when prompted
6. Copy the **Web App URL**
7. In the simulation app, scroll to "Google Sheets Integration" and paste the URL

Each team's submission will create rows in the **Teams** and **Quarters** sheets automatically.

---

## Custom Domain (Optional)

In Vercel dashboard → your project → **Settings → Domains**, you can add a custom domain if you have one.
