# Overview

**This repository is for automating Microsoft Office programs with macros and scripts written in Visual Basic for Applications (VBA) and VBScript**

## Getting Started with VBScript:

1. To create a new VBScript file, open this repository in Visual Studio Code. Add a new file and name it with the .vbs file extension. 
2. When you have written your code and are ready to execute your VBScript, right click the file in the Visual Studio Code Explorer, select "Reveal in File Explorer" (or use keyboard shortcut "SHIFT+ALT+R"), and copy the script file.
3. Navigate to the root directory of your local file system. If you do not have a folder called "Scripts" you should create one. Your path should now look like this: `C:\Users\Your_Username\Scripts>`
4. Open "Command Prompt" by searching in the start menu. Access the "Scripts" folder directory by typing: `cd Scripts`
5. Now type `cscript <YourVBScriptFile.vbs>` so that you should see `C:\Users\Your_Username\Scripts>cscript Example.vbs` and press "Enter" to run your VBScript.

**Note:** In order to successfully interact with an Excel File using VBScript, the Excel File should be located in a folder that is also in your root file directory i.e. `C:\Users\Your_Username\Documents\Test.xlsx`

## VBA for Word & Excel Macros:

1. Create a new Word/Excel File.
2. <ins>For First Time Only:</ins> Enable "Developer" tab in Ribbon by Clicking "File" > "Options" > "Customize Ribbon" > Select the "Developer" checkbox in the "Main Tabs" dropdown menu on the right and click "OK"
3. Navigate back to the "File" > "Options" menu but this time select "Trust Center" and then hit the "Trust Center Settings..." button on the left of the Word Options window.
4. Under "Macro Settings" ensure that the checkbox beneath the "Developer Macro Settings" section is checked. This gives trust access to the VBA project object model and will activate the necessary project references to run a macro. (In Excel, you should also check the box next to the setting titled "Enable Excel 4.0 macros when VBA macros are enabled") Click "OK"
5. Now select the "Developer" tab from the Ribbon in between "View" and "Help" and click the button on the left side of the menu titled "Visual Basic" which will open the Microsoft Visual Basic for Applications Editor.
6. <ins>One Time Only:</ins> In the editor toolbar, click "Tools" > "References..." > search for "Microsoft Word 16.0 Object Library" (In Excel it is called "Microsoft Excel 16.0 Object Library") and check the checkbox. Additionally, ensure that the "Visual Basic for Applications" reference is also checked. Click "OK" in the dialog box to accept your selection(s).
7. Right click on the "Microsoft Word Objects" (or "Microsoft Excel Objects") folder in the Project Explorer menu on the left of the editor. Select "Insert" > "Module" and a new folder titled "Modules" will appear with a macro in it. Double click it and the code editor window should appear. (Alternatively, click "Macros" next to the Visual Basic button in the "Developer" toolbar ribbon, type in a name you want your macro to be called, and click "create" which will then open the VBA Editor with a new module in the "Modules" folder automatically).
8. Place your VBA code in the code editor window and close it when done.
9. Refer to the Project Explorer pane on the left of the VBA Editor and in the "Properties" section you can rename the macro by clicking next to where it says i.e. "(Name) Module1"
10. Now you can close the VBA Editor and click on "Macros" next to the Visual Basic button in the ribbon menu. Select the name of the macro you would like to execute and click "Run"
11. To save your Word/Excel File along with the macro(s), click "File" > "Save As" and select the appropriate file extension when naming your file (.docm is a Microsoft Word macro-enabled document and .xlsm is a Microsoft Excel macro-enabled document).
