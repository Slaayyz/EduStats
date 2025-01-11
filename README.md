# 📊 **EduStats: Academic Performance Analyzer**

**EduStats** is an advanced Python project designed to manage and analyze academic performance for students in the **BUT RT** program. With a user-friendly graphical interface, robust data handling, accurate calculations, and visually intuitive graphs, it offers actionable insights to monitor and enhance results. 📈🎓

---

## ✨ **Key Features**

### **📁 Automated File Management**
- Verifies and organizes academic files (`semester.xlsx`, `notes_RT.xlsx`) automatically.
- Ensures that all required files are present in the `Stats` folder.

### **✅ Enhanced Data Validation**
- Validates the presence and integrity of required Excel files.
- Logs errors and warnings to identify issues efficiently.

### **📊 Weighted Average Calculation**
- Computes averages for each **Teaching Unit (UE)** using module coefficients and user-provided grades.
- Supports flexible grade entry and automatic updates.

### **🎨 Dynamic Visualization**
- Generates clear, color-coded graphs with thresholds for success, retakes, and failure.
- Displays performance results interactively.

### **🖥️ User-Friendly Interface**
- Intuitive GUI for viewing, editing, and saving grades.
- Double-click editing for seamless grade management.

---

## 📂 **File Requirements**

### 1️⃣ **`semester.xlsx`**
- Defines the semester structure with module names and coefficients.
- Must be placed in the `Stats` folder.

### 2️⃣ **`notes_RT.xlsx`**
- Stores grades for each module.
- Automatically created or updated based on `semester.xlsx`.

---

## 🚀 **Installation & Usage**

### 1️⃣ **Install Dependencies**
Ensure Python is installed and required libraries are added:  
```bash
pip install openpyxl matplotlib
```

### 2️⃣ **Clone the Repository**  
```bash
git clone https://github.com/your-username/EduStats.git
cd EduStats
```

### 3️⃣ **Prepare Files**
- Place `semester.xlsx` in the `Stats` folder.
- The program will create or update `notes_RT.xlsx` automatically.

### 4️⃣ **Run the Program**  
```bash
python main.py
```

---

## 🎨 **Graph Overview**

The output graph highlights performance for each **Teaching Unit (UE)** using a color-coded system:  
- **✅ Averages ≥ 10**: Green bars indicate success.  
- **⚠️ Averages between 8 and 10**: Orange bars suggest potential retakes.  
- **❌ Averages < 8**: Red bars flag failure.  

Threshold markers at **10** (pass) and **8** (retake) are included for clarity.

---

## 🛠️ **How It Works**

### **Step 1: File Verification**
- Ensures the `Stats` directory exists.
- Verifies the presence of `semester.xlsx`.
- Creates or updates `notes_RT.xlsx` as needed.

### **Step 2: Data Management**
- Extracts module names and coefficients from `semester.xlsx`.
- Populates `notes_RT.xlsx` for grade entry.

### **Step 3: Data Validation**
- Validates the structure and data of Excel sheets.
- Logs discrepancies for user awareness.

### **Step 4: GUI for Grade Management**
- Allows users to view and edit grades interactively.
- Saves changes to `notes_RT.xlsx` directly.

### **Step 5: Average Calculation and Visualization**
- Computes weighted averages for each **UE**.
- Displays results in an interactive bar chart with performance thresholds.

---

## 💡 **Troubleshooting**

### Common Issues
- **Missing Files**: Ensure `semester.xlsx` is in the `Stats` folder. If `notes_RT.xlsx` is missing, the program will create it.
- **Corrupted Files**: Verify that `semester.xlsx` and `notes_RT.xlsx` are valid Excel files.

### Debugging Tools
- Detailed logging output helps identify and resolve issues efficiently.

If issues persist, check the program output for detailed error messages.

---

## 🛡 **Contributing**  

We welcome contributions! 🚀  
- Report issues or bugs via [GitHub Issues](https://github.com/your-username/EduStats/issues).  
- Submit improvements as pull requests to enhance functionality.  

---

## 📜 **License**  

This project is licensed under the **MIT License**. You are free to use, modify, and distribute it under the terms of this license. Refer to the [LICENSE](./LICENSE) file for details.

---

🎉 **EduStats simplifies academic data management and empowers you to track your success. Start using it today!** 😊

