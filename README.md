# Basic Info

Package Name: **PAteam.BrowserDialog_Actions**

Author: Mateusz Majchrzak

Version: 2.0.0.0

Compatibility: Version 7.5 and higher

# Package Summary

Use this package to handle browser pop-up dialogs such as alert, dialog etc.

Example of browser alert pop-up dialog:  
![enter image description here](https://i.ibb.co/TL4MnFm/clip-image001.png)

Example of browser confirm pop-up dialog:
![enter image description here](https://i.ibb.co/nBK8WjG/clip-image002.png)
  
This package includes multiple functions:

 - Get Active Tab Dialog Existence
 - Get Active Tab Dialog Text
 - Click Active Tab Dialog OK Button
 - Click Active Tab Dialog Cancel Button
 - Open File Dialog On Active Page

**To use this package:**

 - ****RT Designer:****
	 - Download DLL: [DLL](https://github.com/pateam-rpa/NICE-BrowserPopUpActions/blob/master/PAteam.BrowserPopUpActions.Library/bin/Debug/PAteam.BrowserDialog_Actions.dll)
	 - Paste into RT Designer Installation Directory
	 - Add reference from Library References
	 - Start using:
		 - Function are located in Library Objects -> Browser Dialog Actions
- ****Automation Studio****:
	- Download the asset from Resource Center

# Functions

Name: **Get dialog existence**
 - Description: Checks active browser tab for dialog box. If found returns true, else false. Supports only the following languages: PL, EN, DE, SP, IT
 - Parameters:
 - Returns:
   - Description: Active browser tab dialog existence. True  dialog box exists, False  when dialog box doesnt exist or in progress.
   - Type: Boolean

Name: **Get dialog text**
 - Description: Looks at browser active tab to identify dialog box. If found, proceeds to retrieve the dialog body text.
 - Parameters:
 - Returns:
   - Description: Active tab dialog box body text.
   - Type: Text

Name: **Click dialog button**
 - Description: Looks at browser active tab to identify dialog box. If found, proceeds to clicking on button with provided name.
 - Parameters:
   - Description: Dialog button name
   - Type: Text
 - Returns:
   - Description: Status of pressing the OK button. True  button clicked, False  failed to click or still in progress
   - Type: Boolean

Name: **Set dialog edit value**
 - Description: Looks at browser active tab to identify dialog box. If found, proceeds to set provided value in edit value with user provided name.
 - Parameters:
   - Description: Edit element name
   - Type: Text
   - Description: Value to set
   - Type: Text
 - Returns:
   - Description: Status of pressing the Cancel button. True  button clicked, False  failed to click or still in progress.
   - Type: Boolean

Name: **Open file dialog**
 - Description: Looks at browser active tab to identify first file upload button. If found, proceeds to pressing on it, what should result with open file picker window.
 - Parameters:
 - Returns:
   - Description: Status of pressing the upload button. True  button clicked, False  failed to click or still in progress.
   - Type: Boolean