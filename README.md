# AutoNavigationDoc  
### Author: ARGIRO MARIOLI
### GitHub Profile: [p17mari](https://github.com/p17mari)

This repository provides a step-by-step guide on how to add a VBA module to a Word Document that generates a grid listing all files in a folder, including clickable file links.  

## Index  
- [Requirements](#requirements)  
- [Steps to Insert a Grid with File Links](#steps-to-insert-a-grid-with-file-links)
- [Manually add the visual basic code](#manually-add-the-visual-basic-code)
- [Result](#result)  

---

## Requirements  
Before you begin, ensure you have:  
- **Microsoft Word** installed  
- **Macros enabled** in Word  
  - You can enable macros in **Word Options** under **Proofing**  

---

## Steps to Insert a Grid with File Links  
### Run from Template
Follow these steps to generate the navigation document without going over the VBA code

### 1. Download docm file from repository [here](https://github.com/p17mari/AutoNavigationDoc/blob/main/Automatic%20Menu.docm)
### 2. Put Word Document in the folder you want to create the menu for.
### 3. Follow the instructions inside the document:
  - Enable Macros
  - Press the Alt + F8 keys
  - Select InsertFiles
  - Click Run
  - Close Emerging window

## Manually add the visual basic code
Follow these steps to generate a navigation document:  

### 1. Create and Save a Word Document  
- Navigate to the folder where you want to generate the navigation document.  
- Create a new **Microsoft Word document** in this folder.  
- Open the document and **save it** at least once.  

### 2. Open the Visual Basic for Applications (VBA) Editor  
- Press **Alt + F11** to open the **Microsoft Visual Basic for Applications (VBA)** editor.  
  ![image](https://github.com/user-attachments/assets/c9733c4d-0492-449d-8d38-8d1bb5b0c218)  

### 3. Insert a New Module  
- In the **VBA Editor**, go to the top menu, click **Insert**, then select **Module**.  
  ![image](https://github.com/user-attachments/assets/594403f8-538c-45e8-bf61-66e78724320d)  

### 4. Copy and Paste the Code  
- Copy the VBA code from [this file](https://github.com/p17mari/AutoNavigationDoc/blob/main/InsertFileLinks.vb).  
- Paste the copied code into the newly created module.  
- Close the **VBA Editor** window.  

### 5. Run the Macro  
- Return to the Word document.  
- Press **Alt + F8** to open the "Macros" window.  
- Select the macro and click **Run**, then press **OK**.  
  ![image](https://github.com/user-attachments/assets/94ce5553-4dde-4ef4-8eb4-6afb6c493cef)  

---

## Result  
After running the macro, your document will contain a **grid with two columns**:  
- **File Name** – Displays the name of each file or folder  
- **Link** – Provides a clickable link for files  

> **Note:**  
> - If a link is missing, the name belongs to a **folder**.  
> - Files and folders are listed in **alphabetical order**.
> - RecuresiveFileListing recursively lists all files in a folder and its subfolders (no matter how deep)(unsorted)

---

Now your document is ready for easy file navigation! 
