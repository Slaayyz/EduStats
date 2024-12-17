# 📊 EduStats

**EduStats** is a Python project designed to analyze academic data. It reads a CSV file containing grades and credits, calculates weighted averages by semester, and visualizes the results with colorful, intuitive graphs. 🎨📈  

---

## ✨ Features
- **📂 CSV File Parsing**: Automatically imports and processes grade and credit data.
- **📉 Weighted Average Calculation**: Computes averages considering module coefficients.
- **🎨 Data Visualization**: Generates intuitive graphs with thresholds for success and alerts.
- **🗓 Multi-Semester Support**: Seamlessly handles data for multiple semesters.

---

## 📂 Example CSV File
Here’s an example of the `notes.csv` file format:  

```
S1;UE1;R101;10;12
S1;UE1;R102;10;20
S1;UE1;R103;7;13
S2;UE1;R101;10;12
...
```

**Columns**:
- `S1`: Semester.
- `UE1`: Teaching Unit.
- `R101`: Module.
- `10`: Number of Credits.
- `12`: Grade Achieved.

---

## 🚀 Installation and Usage

Follow these steps to set up and run **EduStats**:

### 1️⃣ Clone the Project
```bash
git clone https://github.com/your-username/EduStats.git
cd EduStats
```

### 2️⃣ Install Dependencies
```bash
pip install matplotlib numpy
```

### 3️⃣ Add Your Data
Place a `notes.csv` file in the project directory following the format above.

### 4️⃣ Run the Script
```bash
python script.py
```

---

## 🎨 Graph Overview

The generated graph displays weighted averages for each semester and highlights key thresholds:  
- **✅ Averages >= 10**: Green bars indicate success.  
- **⚠️ Averages between 8 and 10**: Orange bars serve as a warning.  
- **❌ Averages < 8**: Red bars signal failure.  

---

## 🤔 How It Works
1. **Data Parsing**: Reads and processes the `notes.csv` file to extract semester data.  
2. **Weighted Averages**: Calculates averages using grades and their respective credit weights.  
3. **Visualization**: Displays results in a user-friendly graph, highlighting performance thresholds.  

---

## 🛡 Contributing

Contributions are always welcome! If you have ideas or improvements, feel free to open an issue or submit a pull request. 🙌  

---

## 📜 **License**  

This project is open-source and licensed under the **MIT License**. 📄  
You are free to use, modify, and distribute the software under the terms of the license.  

For more information, refer to the [LICENSE](./LICENSE) file.  

---  

🎉 **Enjoy customizing and using EduStats!**  