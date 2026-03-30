# RIPS Excel to JSON Automation

This project automates the conversion of RIPS Excel files into JSON format using a web-based tool. It handles file uploads, downloads, and organization automatically, reducing manual work and improving efficiency.

---

## 🚀 Features

* Automatically detects Excel files (`.xlsx`, `.xlsm`)
* Extracts invoice number from each file
* Uploads files to the SISPRO platform
* Converts Excel files to JSON format
* Downloads generated JSON files
* Organizes outputs into folders by invoice number
* Handles popups and unexpected UI behavior
* Cleans temporary files after processing

---

## 🧰 Technologies Used

* Python
* Selenium
* OpenPyXL
* WebDriver Manager

---

## 📂 Project Structure

```
rips-excel-to-json-automation/
│
├── script.py
├── README.md
├── requirements.txt
```

---

## ⚙️ Configuration

Before running the script, update the following variable in the code:

```python
CARPETA_BASE = r"YOUR_PATH_HERE"
```

This should point to the folder containing your RIPS Excel files.

---

## ▶️ Usage

1. Install dependencies:

```bash
pip install -r requirements.txt
```

2. Run the script:

```bash
python script.py
```

---

## 🔄 Workflow

1. Scan the base folder for Excel files
2. Extract invoice number from each file
3. Upload file to SISPRO platform
4. Convert file to JSON
5. Download the generated JSON
6. Move JSON file to a folder named after the invoice
7. Clean temporary download folder

---

## ⚠️ Requirements

* Google Chrome must be installed
* Stable internet connection
* Access to the SISPRO RIPS conversion website

---

## 📌 Notes

* The script uses browser automation, so the website structure must remain unchanged
* Temporary files are stored in a local folder and deleted after processing
* Errors during processing are logged in the console

---

## 📈 Future Improvements

* Add logging to file
* Improve error handling and retries
* Support batch processing with reports
* GUI interface for easier usage

---

## 👨‍💻 Author

Luis Alejandro Machado.

