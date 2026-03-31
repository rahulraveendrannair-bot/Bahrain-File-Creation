# BH Sanctions List Updater (Offline)

A lightweight Flask web app — **no API key required**.

Upload the Gazette PDF + BH-TL-INDIVIDUALS.xlsx + extracted individuals JSON → download updated XLSX.

---

## Quick Start

```bash
# 1. Install dependencies
pip install flask openpyxl

# 2. Run the app
python app.py

# 3. Open in browser
http://localhost:5000
```

---

## How It Works

1. **Extract individuals** — paste the Gazette PDF into Claude AI chat and ask it to extract individuals as JSON (see `sample_individuals.json` for the format)
2. **Save the JSON** — save Claude's JSON response as a `.json` file
3. **Upload all three files** — PDF + XLSX + JSON in the web app
4. **Download** the updated XLSX

---

## JSON Format

The JSON file must be an array of objects with these exact keys:

```json
[
  {
    "NAME": "Full English name",
    "AKA": "Alias 1; Alias 2",
    "FOREIGN_SCRIPT": "Arabic or non-Latin name",
    "SEX": "Male",
    "DOB": "1 Jul. 1974",
    "POB": "Iraq",
    "NATIONALITY": "Iraqi",
    "OTHER_INFO": "Role, activities, passport/ID numbers...",
    "ADD": "Full address",
    "ADD_COUNTRY": "Iraq",
    "TITLE": "",
    "CITIZENSHIP": "",
    "REMARK": "Gazette issue, UN resolution refs..."
  }
]
```

See `sample_individuals.json` for a complete example.

---

## Deploy to GitHub + Render (free hosting)

### Step 1 — Push to GitHub
```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/sanctions-updater.git
git push -u origin main
```

### Step 2 — Deploy on Render (free)
1. Go to https://render.com and sign up free
2. Click **New** → **Web Service**
3. Connect your GitHub repo
4. Set these values:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
5. Click **Deploy**

Add `gunicorn` to `requirements.txt` for Render:
```
flask>=3.0.0
openpyxl>=3.1.0
gunicorn>=21.0.0
```

Your app will be live at `https://your-app-name.onrender.com`

---

## Files

```
sanctions-offline/
├── app.py                   ← Flask web app (all-in-one)
├── requirements.txt         ← Python dependencies
├── sample_individuals.json  ← Example JSON format
└── README.md
```
