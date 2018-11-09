# LogicAndSetSymbols
Custom logic and set symbols menu for MS Word

This is a basic menu I created to easily insert logic and set symbols in MS Word:

[GIF Example](https://i.imgur.com/qoi1XTd.gifv)

## Instructions

### 1. Enable Developer Console in Word

1. Click on the “File” tab and select “Options.” 
2. Click on “Customize Ribbon.”
3. Select “Main Tabs” from the dropdown menu below “Customize the Ribbon.”
4. Place a checkmark next to “Developer.”
5. Click on “OK.”

### 2. Import logic symbols menu

1. In the "Developer" tab, select "Visual Basic"
2. In the Project pane (top left corner), expand "Normal"
3. Right click on the "Forms" folder, and select "Import File..."
4. Select "LogicSymbolsNew.frm" and press okay.

### 3. Add a macro to open the logic symbols menu

1. In the Project pane of the Visual Basic window, right click on "Modules" and select Insert > Module.
2. Paste in the below code:

```
Sub LogicSymbolsShow()
'
' LogicSymbolsShow Macro
'
'
LogicSymbolsNew.Show vbModeless
End Sub
```

3. Close the VB window, saving if prompted.

### 4. Add a button to the Quick Access Toolbar to open the menu

1. Right click on the ribbon and select "Customize Quick Access Toolbar..."
2. On the top-left menu (below "Choose commands from:") select "Macros"
3. Select "Normal.NewMacros.LogicSymbolsShow" and click "Add >>" to move it to the righthand pane
4. You can customize the button's appearance by selecting "Normal.NewMacros.LogicSymbolsShow" in the righthand menu
and pressing the "Modify..." button below.
