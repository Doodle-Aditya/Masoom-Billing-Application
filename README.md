# Master on Billing Application

## 📌 Overview
**Master on Billing Application** is a desktop-based billing system developed for the organization **Masoom**. The application simplifies the billing workflow by allowing billing staff to enter billing details through a user-friendly GUI, automatically generate Excel-based bills, maintain a centralized billing record, and convert bills into PDF format for storage and sharing.

This project is built using **Python** with a structured folder architecture to ensure scalability, maintainability, and clarity.

---

## 🚀 Features
- Intuitive GUI for billing staff using **CustomTkinter**
- Automatic filling of billing details into a predefined Excel bill template
- Centralized billing record maintenance (`record.xlsx`)
- Automatic conversion of Excel bills into PDF format
- Organized output storage for generated Excel files and PDFs
- Clean and modular code structure

---

## 📂 Project Structure

    Master on Billing Application/
    │
    ├── Excel Handler/
    │   ├── ExcelManager.py        # Handles Excel read/write operations
    │   └── FillTemplate.py        # Fills the bill template with user input
    │
    ├── Excel Template/
    │   └── BillTemplate.xlsx      # Predefined bill template
    │
    ├── GUI/
    │   └── (GUI-related Python files)
    │       # All UI logic and layouts
    │
    ├── Output/
    │   ├── record.xlsx            # Central billing record file
    │   └── (Generated bills)
    │       # Excel and PDF bills are stored here
    │
    ├── main.py                    # Entry point to run the application
    ├── requirements.txt           # Project dependencies
    └── README.md                  # Project documentation

---

## 🧾 How the Application Works
1. The billing staff enters billing details through the GUI.
2. The application:
   - Fills the entered data into the **BillTemplate.xlsx** file.
   - Updates a centralized record file (`record.xlsx`) containing:
     - Customer/Student name
     - Billing person name
     - Billing amount
     - Masoom's contribution
     - Other relevant billing details
3. The completed Excel bill is automatically converted into a **PDF**.
4. Both the Excel file and the generated PDF are saved in the **Output** folder.

---

## 🛠️ Technologies & Libraries Used
- **Python**
- **CustomTkinter** – for modern GUI design
- **OpenPyXL** – for Excel file handling
- **pywin32** – for Excel to PDF conversion
- **num2words** – for converting numeric amounts into words

---

## 📦 Installation

### 1️⃣ Clone the Repository
    git clone https://github.com/your-username/master-on-billing-application.git
    cd master-on-billing-application

### 2️⃣ Install Dependencies
    pip install -r requirements.txt

> ⚠️ **Note:** `pywin32` requires Microsoft Excel to be installed on Windows.

---

## ▶️ Running the Application
    python main.py

The GUI will launch, allowing billing staff to start generating bills.

---

## 🏢 Organization
This application was developed specifically for **Masoom** to streamline and digitize their billing operations.

---

## 📬 Contact
For any questions or suggestions, feel free to reach out:

- 📧 Email: adityanishad98196@gmail.com
- 💼 LinkedIn: https://www.linkedin.com/in/aditya-nishad-938403330/

---

⭐ If you find this project useful, consider giving it a star!
