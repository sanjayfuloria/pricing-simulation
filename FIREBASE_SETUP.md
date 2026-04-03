# Firebase Setup Guide — Pricing Simulation

## Why Firebase?
Firebase Realtime Database provides **free** real-time sync between faculty and student browsers.
When the faculty starts a quarter, ALL students see it instantly. When students submit prices,
the faculty dashboard updates in real-time.

**Free tier includes:** 1GB storage, 10GB/month transfer, 100 simultaneous connections.
**No credit card required.**

---

## Step 1: Create a Firebase Project (2 minutes)

1. Go to **https://console.firebase.google.com**
2. Click **"Create a project"** (or "Add project")
3. Name it: `pricing-simulation`
4. Disable Google Analytics (not needed) → **Create project**
5. Wait for creation → **Continue**

## Step 2: Add a Web App

1. On the project overview page, click the **Web icon** `</>`
2. App nickname: `pricing-sim-web`
3. **Don't** check "Firebase Hosting"
4. Click **Register app**
5. You'll see a `firebaseConfig` object — **copy these values**:
   ```javascript
   const firebaseConfig = {
     apiKey: "AIza...",
     authDomain: "pricing-simulation-xxxxx.firebaseapp.com",
     databaseURL: "https://pricing-simulation-xxxxx-default-rtdb.firebaseio.com",
     projectId: "pricing-simulation-xxxxx",
     storageBucket: "pricing-simulation-xxxxx.appspot.com",
     messagingSenderId: "123456789",
     appId: "1:123456789:web:abcdef..."
   };
   ```
6. Click **Continue to console**

## Step 3: Enable Realtime Database

1. In the left sidebar, click **Build → Realtime Database**
2. Click **Create Database**
3. Location: **United States (us-central1)** or closest to you
4. Security rules: Select **Start in test mode** → **Enable**

   > Note: Test mode allows anyone to read/write for 30 days. 
   > You should update rules later for production use.

## Step 4: Add Config to Vercel

1. Go to **https://vercel.com** → your `pricing-simulation` project
2. Click **Settings → Environment Variables**
3. Add each of these (copy values from Step 2):

   | Key | Value |
   |-----|-------|
   | `VITE_FIREBASE_API_KEY` | `AIza...` |
   | `VITE_FIREBASE_AUTH_DOMAIN` | `pricing-simulation-xxxxx.firebaseapp.com` |
   | `VITE_FIREBASE_DATABASE_URL` | `https://pricing-simulation-xxxxx-default-rtdb.firebaseio.com` |
   | `VITE_FIREBASE_PROJECT_ID` | `pricing-simulation-xxxxx` |
   | `VITE_FIREBASE_STORAGE_BUCKET` | `pricing-simulation-xxxxx.appspot.com` |
   | `VITE_FIREBASE_MESSAGING_SENDER_ID` | `123456789` |
   | `VITE_FIREBASE_APP_ID` | `1:123456789:web:abcdef...` |

4. Click **Save** for each
5. **Redeploy**: Go to Deployments tab → click "..." on latest → Redeploy

## Step 5: Test

1. Open your app in **two browser tabs**
2. Tab 1: Faculty → Create game → Set up teams → Start Quarter 1
3. Tab 2: Student → Enter game code → Select team → Submit price
4. Faculty tab should show the submission in real-time!

## Local Development

For local testing, create a `.env` file in the project root:
```
VITE_FIREBASE_API_KEY=AIza...
VITE_FIREBASE_AUTH_DOMAIN=...
# (copy all values from .env.example)
```

Then run `npm run dev`.

## Security Rules (for production)

After testing, update your Realtime Database rules:
```json
{
  "rules": {
    "games": {
      "$gameId": {
        ".read": true,
        "meta": {
          ".write": true
        },
        "teams": {
          "$teamId": {
            ".write": true
          }
        }
      }
    }
  }
}
```

This allows anyone to read game data but only write to specific paths.
