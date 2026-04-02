# Pricing Simulation — B2B Marketing, IBS Hyderabad

A competitive pricing simulation game for MBA students. Teams set prices for government tiffin centre meals across 12 quarters (3 phases), competing for customers and profit.

## Game Overview

### Phase 1 — Market Entry (Quarters 1–4)
- Variable Cost: ₹3/meal | Fixed Cost: ₹500/quarter
- New Customers = 400 − 40×(Own Price) + 21×(Avg Competitor Price)

### Phase 2 — Promotions (Quarters 5–8)
- Same costs as Phase 1
- New Customers = 400 − 40×(Own Price) + 21×(Avg Comp Price) + 0.10×(Promo Expense)
- Promotion budget capped at 20% of last quarter revenue

### Phase 3 — Recession (Quarters 9–12)
- Variable Cost: ₹5/meal | Fixed Cost: ₹1000/quarter
- New Customers = 1000 − 40×(Own Price) + 21×(Avg Comp Price) + 0.20×(Promo Expense) + 100×(Last Qtr Price − This Qtr Price)

### Customer Retention (all phases)
| Your Price vs Last AIP | Retention Rate |
|------------------------|---------------|
| More than 20% lower    | 30%           |
| 10–20% lower           | 15%           |
| Less than 10% lower    | 10%           |
| ≥ AIP                  | 0%            |

---

## Deploy to Vercel

### Prerequisites
- Node.js 18+ installed
- A [Vercel account](https://vercel.com)

### Steps

```bash
# 1. Navigate to the project folder
cd pricing-simulation

# 2. Install dependencies
npm install

# 3. Test locally
npm run dev

# 4. Deploy to Vercel (one of these):

# Option A — Vercel CLI
npm i -g vercel
vercel

# Option B — Push to GitHub, then import on vercel.com
```

---

## Google Sheets Integration

### Step 1: Create a Google Sheet
Create a new spreadsheet in Google Sheets. Name it anything you like.

### Step 2: Add the Apps Script
1. Open the sheet → **Extensions → Apps Script**
2. Delete any default code
3. Copy the contents of `google-apps-script.js` and paste it
4. Save (Ctrl+S)

### Step 3: Deploy the Script
1. Click **Deploy → New deployment**
2. Click the gear icon → select **Web app**
3. Set:
   - **Execute as:** Me
   - **Who has access:** Anyone
4. Click **Deploy**
5. **Authorize** when prompted (grant Sheets access)
6. Copy the **Web App URL**

### Step 4: Connect the App
Paste the Web App URL in the "Google Sheets Integration" section of the simulation app.

### What Gets Saved
- **Teams sheet**: One row per team submission with totals
- **Quarters sheet**: One row per quarter per team with full detail

---

## Project Structure

```
pricing-simulation/
├── index.html
├── package.json
├── vite.config.js
├── vercel.json
├── google-apps-script.js    ← Paste into Google Apps Script
├── README.md
└── src/
    ├── main.jsx
    ├── App.jsx               ← All simulation logic
    └── index.css
```

## Running the Simulation in Class

1. Deploy the app to Vercel
2. Share the URL with all teams
3. Each team fills in their name, members, and strategy
4. The instructor announces the Average Industry Price (AIP) each quarter
5. Teams set their price and submit
6. All data flows to the shared Google Sheet for the instructor to review

---

*Built for the B2B Marketing course, MBA 2026, IBS Hyderabad.*
