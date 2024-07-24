# Overview

**This repository is for automating Microsoft Office programs with macros and scripts written in Visual Basic for Applications (VBA) and VBScript**

## Getting Started with VBScript:

1. To create a new VBScript file, open this repository in Visual Studio Code. Add a new file and name it with the .vbs file extension. 
2. When you have written your code and are ready to execute your VBScript, right click the file in the Visual Studio Code Explorer, select "Reveal in File Explorer" (or use keyboard shortcut "SHIFT+ALT+R"), and copy the script file.
3. Navigate to the root directory of your local file system. If you do not have a folder called "Scripts" you should create one. Your path should now look like this: `C:\Users\Your_Username\Scripts>`
4. Open "Command Prompt" by searching in the start menu. Access the "Scripts" folder directory by typing: `cd Scripts`
5. Now type `cscript <YourVBScriptFile.vbs>` so that you should see `C:\Users\Your_Username\Scripts>cscript Example.vbs` and press "Enter" to run your VBScript.

**Note:** In order to successfully interact with an Excel File using VBScript, the Excel File should be located in a folder that is also in your root file directory i.e. `C:\Users\Your_Username\Documents\Test.xlsx`
