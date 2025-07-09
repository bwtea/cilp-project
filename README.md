# cilp-project
Just a small program to filter and combine exel files.

# 🧩 CIPL — Excel Assistant Tool

A lightweight tool to process Excel files in 4 simple steps:

---

## 🔧 What does it do?

1. **Filters rows from Excel files**
   - Starts from column **E**
   - Skips empty rows
   - Ignores column **G** (a formula will be added later)
   - Works only with the worksheet named `CIPL`

2. **Combines filtered tables**
   - Stacks all filtered tables into one
   - Removes any rows where column **Q** is empty
   - Adds a `Total` row at the end:
     - Sum of column **D** → into column **D**
     - Word `"Total"` → into column **E**
     - Sum of column **F** → into column **F**

3. **Attaches the final table to a template**
   - Uses a built-in Excel template (`Template.xlsx`)
   - Inserts the table starting at **column I, row 13**
   - Columns **A to H** remain empty but retain styling
   - All rows use the style of row **13**
   - The final `Total` row uses the style of row **29**

---

## ✅ How to use it

1. Click **`Open Exel Files`**  
   👉 Select one or  `.xlsx` file to filter.
   
2. Click **`Filter and Save`**  
   👉 Save and filter previously selected file.
   
   **Do it with all Exel Files you got. 
   
3. Click **`Merge Multiple Excel Files`**  
   👉 Select one or more `.xlsx` filtred files .

4. Click **`Attach to Final Template`**  
   👉 This attaches the merged result to the built-in template and saves the final report.

---

## 📦 What’s included?

- `CIPL.exe` — the main application
- `Assets/Template.xlsx` — the built-in Excel template
- `app.ico` — the application icon
- `README.md` — this guide

---

## 📌 Requirements

- Windows 10 or 11
- Excel `.xlsx` files
- [.NET 8+ Runtime](https://dotnet.microsoft.com/en-us/download)

---

## 💚 Created by

🛠 Developed by **Natalie**, with care and attention to detail for NY INTERNATIONAL LOGISTICS CO., LTD.
📍 Seoul, 2025
