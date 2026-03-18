# 🏫 Attendance Defaulter System
**Pimpri Chinchwad College of Engineering — Computer Engineering**
**Live - https://attendance-system-tau-liart.vercel.app/**

Upload the attendance Excel sheet → auto-detects defaulters (<75%) → generate letters as PDF or DOCX.

---

## 📁 Project Structure
```
attendance-system/
├── server/
│   ├── index.js          ← Express backend
│   ├── template.html     ← PDF letter template
│   ├── logo.png          ← College logo (copy from ref letter)
│   ├── uploads/          ← Temp Excel uploads (auto-created)
│   ├── letters/          ← Generated letters (auto-created)
│   └── package.json
└── client/
    ├── src/
    │   ├── App.js
    │   └── index.js
    ├── public/index.html
    └── package.json
```

---

## ⚙️ Setup & Run

### 1️⃣ Backend Setup
```bash
cd server
npm install
node index.js
```
Server starts on → **http://localhost:5000**

> ⚠️ First time puppeteer install may take a few minutes (downloads Chrome)

### 2️⃣ Frontend Setup (new terminal)
```bash
cd client
npm install
npm start
```
App opens at → **http://localhost:3000**

---

## 🚀 How to Use

1. Open **http://localhost:3000**
2. Click **Choose Excel File** → select your attendance `.xlsx`
3. Click **Find Defaulters** → table loads all students below 75%
4. For each student, click:
   - 📄 **PDF** → downloads a PDF letter
   - 📝 **DOCX** → downloads a Word document letter
5. Or use **Download All PDF / Download All DOCX** to bulk-download

---

## 📊 Excel Format Expected

The system auto-detects subjects from your Excel header row.
It reads the row containing `PRN` and `Name of the Student` as headers.

Columns it looks for:
- `PRN` — student PRN
- `Name of the Student` — student name  
- `SUBJECT-TH` — theory attendance (e.g. `CC-TH`, `STQA-TH`)
- `SUBJECT-PR` or `SUBJECT-L-PR` — practical attendance (e.g. `CCL-PR`, `STQAL-PR`)
- `Overall TH Att.` — overall theory %
- `Overall PR Att.` — overall practical %
- `Overall Att.` — total overall % (used to filter <75%)

---

## ✉️ Letter Format

- College letterhead with logo (PCCOE format)
- Student name, PRN, division, semester
- Subject-wise attendance table (auto-extracted from Excel headers)
- Average TH / PR attendance row
- Total attendance row
- Signatures: Class Teacher, Academic Coordinator, HOD
- **No date** 

---

## 🔧 Customization

### Change signature names


### Change 75% threshold
In `server/index.js`, find:
```js
if (overall < 75) {
```
Change `75` to your desired cutoff.

### Change logo
Replace `server/logo.png` with your college logo (PNG, ~200x200px recommended).
