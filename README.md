# 🏡 Home Maintenance Logger

A desktop application to log and manage home maintenance tasks.  
Built with **Python + PySide6** for a modern UI.  

---

## ✨ Features
- 📋 **CRUD**: Add, Update, Delete, and Search records  
- 💾 **Auto-save** to `data.json` (local storage)  
- 📂 **Export to Excel** (`data_YYYY-MM-DD.xlsx`)  
- 📥 **Import JSON** backups  
- 🎨 **Modern UI** (PySide6 + optional `qt-material` themes)  
- 🖥 **Always centered window** for better usability  
- ✅ No login required — lightweight and easy to use  

---

## 🚀 Installation & Usage

### 1. Clone the repository
```bash
git clone https://github.com/your-username/home-maintenance-logger.git
cd home-maintenance-logger
````

### 2. Create a virtual environment (optional but recommended)

```bash
python -m venv venv
source venv/bin/activate   # Mac/Linux
venv\Scripts\activate      # Windows
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

If you don’t have a `requirements.txt` yet, install manually:

```bash
pip install PySide6 pandas qt-material
```

### 4. Run the program

```bash
python main.py
```

---

## 📦 Data Storage

* Data is stored in `data.json` automatically.
* Exported Excel files are named like `home_maintenance_2025-09-24.xlsx`.

---

## 🔑 Requirements

* Python **3.9+**
* [PySide6](https://pypi.org/project/PySide6/)
* [pandas](https://pypi.org/project/pandas/)
* (Optional) [qt-material](https://pypi.org/project/qt-material/) for themes

---

## 📘 Example Use Case

1. Log water heater servicing under **Plumbing** with cost and next service date.
2. Track AC filter replacement under **HVAC**.
3. Export all records monthly to Excel for household budgeting.

---

## 📄 License

This project is licensed under the MIT License.
Feel free to use and modify for personal or professional projects.

---

## 🙌 Acknowledgements

* [PySide6](https://doc.qt.io/qtforpython/) for GUI framework
* [qt-material](https://github.com/UN-GCPDS/qt-material) for modern themes
* [pandas](https://pandas.pydata.org/) for Excel export support
