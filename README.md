# m365-office-macros

**A curated collection of Microsoft 365 Office macros (Excel, Word, Outlook, PowerPoint) designed to streamline workflows, automate repetitive tasks, and boost productivity.**

## 📌 About

This repository is a central hub for handy Microsoft Office macros you can plug into your Microsoft 365 environment. Whether you’re looking to clean up data in Excel, automate repetitive formatting in Word, manage emails in Outlook, or speed up slide design in PowerPoint — this is where they live.

Each macro is:

* **Tested** in Microsoft 365
* **Documented** with usage instructions
* **Organized** by application

## 📂 Structure

```
/Excel
  MacroName.bas
  README.md
/Word
  MacroName.bas
  README.md
/Outlook
  MacroName.bas
  README.md
/PowerPoint
  MacroName.bas
  README.md
```

## 🚀 Getting Started

### 1. Download the Macro

* Navigate to the folder for your Office app
* Download the `.bas` file you want

### 2. Import into Office

* Open the VBA editor (`Alt + F11`)
* Go to **File → Import File…**
* Select the `.bas` file
* Save and run the macro

### 3. Enable Macros in Microsoft 365

Make sure your **Macro Security Settings** allow you to run macros.

* Go to **File → Options → Trust Center → Trust Center Settings → Macro Settings**
* Select the appropriate option for your security needs

## 🛠 Contributing

Got a macro that saves time? Submit a pull request! Please include:

* The `.bas` file
* A short description of what it does
* Any dependencies or special setup notes

## 📜 License

This project is licensed under the GNU General Public License v3.0&#x20;
