# 📊 **EduStats: Academic Performance Analyzer**

**EduStats** is an advanced Python project tailored to manage and analyze academic performance for students in the **BUT RT** program. With seamless Excel file handling, precise calculations, and visually intuitive graphs, it provides actionable insights to help track and improve results. 📈🎓  

---

## ✨ **Key Features**
- **📁 Automated File Management**  
  - Verifies and organizes academic files (`semestre.xlsx`, `notes_RT.xlsx`) automatically.
  - Creates or updates the `notes_RT.xlsx` file based on the semester structure.  

- **📊 Weighted Average Calculation**  
  - Computes averages for each **Teaching Unit (UE)** using module coefficients and user-provided grades.

- **🎨 Dynamic Visualization**  
  - Generates clear, color-coded graphs with performance thresholds for success, warnings, and failures.

- **🛠 Easy-to-Use Workflow**  
  - Intuitive setup ensures a smooth experience from file preparation to result visualization.

---

## 📂 **File Requirements**

### 1️⃣ **`semestre.xlsx`**  
- Defines the semester structure with module names and coefficients.  
- Must be placed in the `Stats` folder.  

### 2️⃣ **`notes_RT.xlsx`**  
- Stores the grades for each module.  
- Created or updated automatically based on `semestre.xlsx`.

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
- Place `semestre.xlsx` in the `Stats` folder.  
- The program will create or update `notes_RT.xlsx` for you.  

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
- Verifies the presence of `semestre.xlsx`.  
- Creates or updates `notes_RT.xlsx` automatically.

### **Step 2: Data Management**  
- Extracts module names and coefficients from `semestre.xlsx`.  
- Populates `notes_RT.xlsx` with a user-friendly structure for grade input.  

### **Step 3: Average Calculation**  
- Computes weighted averages for each **UE**, accounting for module coefficients and grades.

### **Step 4: Visual Results**  
- Displays a visually intuitive bar chart with performance thresholds and color-coded bars.  

---

## 💡 **Troubleshooting**
- **Missing Files**: Ensure `semestre.xlsx` is in the `Stats` folder. If `notes_RT.xlsx` is missing, the program will create it.  
- **Corrupted Files**: Verify that `semestre.xlsx` and `notes_RT.xlsx` are valid Excel files.  

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