# LogicAndSetSymbols
Custom logic and set symbols menu for MS Word

This is a basic menu I created to easily insert logic and set symbols in MS Word:

[GIF Example](https://i.imgur.com/qoi1XTd.gifv)

## Instructions

### 1. Enable Developer Console in Word

-Click on the “File” tab and select “Options.” 

-Click on “Customize Ribbon.”

-Select “Main Tabs” from the dropdown menu below “Customize the Ribbon.”

-Place a checkmark next to “Developer.”

-Click on “OK.”

### 2. Import logic symbols menu

-In the "Developer" tab, select "Visual Basic"

-In the Project pane (top left corner), expand "Normal"

-Right click on the "Forms" folder, and select "Import File..."

-Select "LogicSymbolsNew.frm" and press okay.

### 3. Add a macro to open the logic symbols menu

-In the Project pane of the Visual Basic window, right click on "Modules" and select Insert > Module.

-Paste in the below code:



Sub LogicSymbolsShow()
'
' LogicSymbolsShow Macro
'
'
LogicSymbolsNew.Show vbModeless

End Sub



-Close the VB window, saving if prompted.

### 4. Add a button to the Quick Access Toolbar to open the menu

-Right click on the ribbon and select "Customize Quick Access Toolbar..."

-On the top-left menu (below "Choose commands from:") select "Macros"

-Select "Normal.NewMacros.LogicSymbolsShow" and click "Add >>"

-You can customize the button's appearance by selecting "Normal.NewMacros.LogicSymbolsShow" in the righthand menu
and pressing the "Modify..." button below.
