# vba-pharma-delivery-acceptance

A reusable VBA macro to automate the generation of pharmaceutical delivery acceptance reports using **Word mail merge templates** and **Excel data sources**.

This tool was designed to replace manual processes with a single-click batch operation, enhancing data accuracy and reducing report generation time from 5–20 minutes per unit to under 3 minutes for 20+ units.

---

## 📌 Project Background

In many hospitals or pharmacy settings, staff are required to generate daily delivery acceptance records based on supplier shipments. Originally, this process was manually performed for each unit:

- Open Excel delivery data  
- Copy relevant data to a Word template  
- Generate and save as DOC and PDF files  

This macro automates that entire process by:
- Identifying rows based on pharmacy name (keyword)
- Merging Excel data into a predefined Word mail merge template
- Saving output reports with standardized filenames in both DOC and PDF formats

---

## 🛠 Features

- Batch automation for 20+ units
- Search-and-match keyword-based row filtering
- Mail Merge execution in Word
- Save output in `.doc` and `.pdf` formats
- Fully modular and reusable design

---

## 💡 How It Works

1. Excel data should contain delivery records with a pharmacy unit name (used as keyword).
2. The macro searches the sheet for a matching row.
3. It opens the corresponding Word mail merge template.
4. Executes the mail merge and saves the output.

---

## 📁 Repository Structure
/src/
    DeliveryMergeModule.bas        ← VBA macro module


---

## ▶️ How to Use

1. Open Excel, press `Alt + F11` to launch the VBA editor.  
2. Import `DeliveryMergeModule.bas` into a module.  
3. Modify the `Main_DeliveryMerge` subroutine with your file paths and keywords.  
4. Run the macro to batch generate reports.  

---

## 🚫 No Sample Data

To respect data confidentiality, no actual Excel or Word files are included.  
However, you can prepare your own based on these field suggestions:

**Excel Required Columns:**
- Pharmacy Unit Name (e.g., `Unit01`, `Unit02`)
- Product Name
- Quantity
- Receiver
- Remarks

**Word Template Requirements:**
- Set up with mail merge fields like: `<<Pharmacy>>`, `<<Product>>`, `<<Quantity>>`, etc.

---

## 🧩 License

MIT License – Free for use and modification. No warranties provided.

---

## ✉️ Contact

Created by JessieSeal0122  
Feel free to connect on [GitHub](https://github.com/JessieSeal0122)
