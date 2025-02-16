# 📊 Excel to PowerPoint Generator Web App

A Flask web application that processes uploaded Excel files and generates PowerPoint presentations based on the provided data.

## 📂 Project Structure
```
project-folder/
│   app.py                # Flask application (routes)
│   excel_to_ppt.py       # Module to process Excel and create PowerPoint
│   requirements.txt      # Python dependencies
│   README.md             # Project documentation
└───templates/
└───static/
└───output/              # Output folder for generated PPT files
```

---

## ⚙️ Installation

### 1. **Clone the repository:**
```bash
git clone https://github.com/yourusername/excel-to-ppt-flask.git
cd excel-to-ppt-flask
```

### 2. **Create a virtual environment:**
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

### 3. **Install dependencies:**
```bash
pip install -r requirements.txt
```

---

## 🚀 Running the App
```bash
flask run
```
- Open your browser and visit: `http://127.0.0.1:5000/`

If using a different port:
```bash
flask run --port=8080
```

---

## 🛑 Troubleshooting
- **PermissionError:** Ensure the `output/` directory has proper write permissions.
  ```bash
  mkdir output
  chmod 777 output
  ```
- **ModuleNotFoundError:** Ensure `__init__.py` files exist for all modules.

---

## ✅ Example Workflow
1. Upload an Excel file via the web interface.
2. The app processes the file and creates a PowerPoint presentation.
3. Download the generated `.pptx` file.

---

## 📌 Requirements
- Python 3.12.8
- Flask
- python-pptx (for creating PowerPoint)
- pandas (for processing tabular data)

---

## 📝 License


