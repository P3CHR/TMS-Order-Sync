# 📊 TMS Order Sync

> Automated order synchronization between **TMS back-office** and a local **Excel tracker**.  
> Keeps your sales/operations workflow fast, accurate, and low-maintenance.

---

## 🌍 Why this project exists

Managing dozens (sometimes hundreds) of orders in TMS every week is:

- ⏳ **Time-consuming** – endless clicks and manual updates  
- ⚠️ **Error-prone** – missing receipts, mismatched purchases, overlooked interruptions  
- 🔄 **Inefficient** – double work between the system and Excel  

**TMS Order Sync** solves this by automating the entire loop: fetch → validate → update.  
The result: an always-accurate Excel tracker without the manual pain.

---

## ⚙️ What it does

✅ Logs in automatically to the TMS system (session-based, token handled)  
✅ Fetches all orders linked to a given sales user  
✅ Parses order details (status, payments, customer, items, purchases…)  
✅ Flags issues like missing receipts or inconsistent purchases  
✅ Inserts new orders only once, updates existing ones automatically  
✅ Updates a local Excel tracker for team visibility  

---

## 🚀 Why it’s useful

- ⏱ **Saves hours per week** of repetitive admin  
- 🧾 **Reduces mistakes** by letting code catch inconsistencies  
- 📑 **Keeps Excel as the single source of truth** (familiar to all staff)  

---

## 🔧 Tech stack

- [Python 3.11+](https://www.python.org/)  
- [requests](https://pypi.org/project/requests/) – HTTP sessions  
- [beautifulsoup4](https://pypi.org/project/beautifulsoup4/) – HTML parsing  
- [openpyxl](https://pypi.org/project/openpyxl/) – Excel automation  

---
