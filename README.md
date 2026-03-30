# Shopify Milestone Creator

A smart milestone planning tool built using **Streamlit** to generate structured Shopify project timelines with dynamic task allocation and export-ready reports.

---

## ✨ Features

### 📋 Dynamic Milestone Planning
- Auto-generates milestones (M1 → M8)
- Adjusts tasks based on:
  - Client type (New / Retainer)
  - Design approach (Figma / Direct)
  - Custom development
  - SEO, App integration, Payment gateway

### ⚙️ Smart Configuration
- Flexible timeline:
  - Start Date + Working Days
  - Start Date + End Date
- Working-day logic (weekends excluded)
- Dynamic scaling of milestone durations

### ➕ Custom Add-ons
- Add custom tasks to any milestone
- Automatic timeline recalculation
- Visual milestone reference system

### 🎛️ Interactive Adjustments
- Modify milestone duration post-generation
- Automatic cascading date recalculation

### 📊 Export Options
- 📄 Professional PDF report (client-ready)
- 📊 Structured Excel sheet (editable)

### 🎨 Clean UI
- Dark themed UI with intuitive workflow
- Visual milestone preview cards
- Summary dashboard

---

## 🧠 Use Case

Perfect for:
- Shopify Agencies
- Project Coordinators
- Freelancers
- Development Teams

Helps:
- Standardize project planning
- Improve client communication
- Reduce manual timeline errors

---

## 🛠️ Tech Stack

- **Frontend/UI:** Streamlit
- **Backend Logic:** Python
- **PDF Generation:** ReportLab
- **Excel Export:** OpenPyXL

---

## ⚡ Installation

```bash
git clone https://github.com/your-username/shopify-milestone-creator.git
cd shopify-milestone-creator
pip install -r requirements.txt
streamlit run app.py
