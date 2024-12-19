# 📊 EduStats

**EduStats** is a Python project designed to manage and analyze academic data for a BUT RT program. It automates the creation and maintenance of necessary Excel files, calculates averages per teaching unit (UE), and visualizes the results through intuitive graphs. 🔢🔯

---

## ✨ Features
- **📂 Directory and File Validation**: Ensures the required files (`semestre.xlsx`, `notes_RT.xlsx`) exist, creating or updating them if necessary.
- **📊 Automatic File Management**: Extracts and organizes data from `semestre.xlsx` into `notes_RT.xlsx`.
- **🔢 Average Calculation**: Computes averages for each UE based on user-provided notes and module coefficients.
- **🎨 Visual Results**: Displays results in a graph with clear thresholds for success and alerts.

---

## 📂 Required Files
### 1. **`semestre.xlsx`**
Contains the structure of semesters and their corresponding modules with coefficients.
### 2. **`notes_RT.xlsx`**
Generated or updated automatically to allow users to input grades for each module.

---

## 🚀 Installation and Usage

### 1️⃣ Prerequisites
Ensure Python is installed on your system, along with the required libraries.

```bash
pip install openpyxl matplotlib
```

### 2️⃣ Clone the Project
```bash
git clone https://github.com/your-username/EduStats.git
cd EduStats
```

### 3️⃣ Place the Files
Place your `semestre.xlsx` file in the `Stats` folder. The program will handle the rest.

### 4️⃣ Run the Script
```bash
python script.py
```

---

## 🎨 Graph Overview

The generated graph visually represents the averages for each UE, using a color-coded system:
- **✅ Averages >= 10**: Green bars indicate success.
- **⚠️ Averages between 8 and 10**: Orange bars warn of retakes.
- **❌ Averages < 8**: Red bars signal failure.

Threshold lines for 10 (pass) and 8 (retake) are also displayed for clarity.

---

## 🤔 How It Works

### 1. **Directory and File Checks**
- Verifies the existence of the `Stats` directory.
- Checks for `semestre.xlsx` and ensures `notes_RT.xlsx` is created or updated.

### 2. **Automatic File Management**
- **File Creation**: Extracts module titles and coefficients from `semestre.xlsx` to initialize `notes_RT.xlsx`.
- **File Update**: Adds missing modules to `notes_RT.xlsx` if new entries appear in `semestre.xlsx`.

### 3. **Average Calculation**
- Reads grades and coefficients from `notes_RT.xlsx`.
- Computes weighted averages for each UE.

### 4. **Graph Display**
- Displays a bar chart with averages, color-coded by performance thresholds.

---

## 🛡️ Troubleshooting
- Ensure both `semestre.xlsx` and `notes_RT.xlsx` are placed in the `Stats` folder.
- If files are corrupted or missing, the program will provide guidance or exit gracefully.

---

## 🛠 Contributing

Contributions are welcome! Feel free to submit issues or pull requests to improve **EduStats**. 🙏

---

## 📜 License

This project is open-source and licensed under the **MIT License**. You are free to use, modify, and distribute the software under the terms of the license. Refer to the [LICENSE](./LICENSE) file for more information.

---

🎉 **Enjoy automating your academic data management with EduStats!**

