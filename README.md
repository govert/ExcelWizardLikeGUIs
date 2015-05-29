# ExcelWizardLikeGUIs
Research on how to build GUIs which can select ranges into Excel, within an Excel add-in

Note
--------
This project is for research only, it is not meant to be an usable tool in the end.

Overview
--------
The purpose is to find out how to make a GUI available from an Excel add-in, which will allow to select a range in order to populate the address into some control within that GUI.

- The GUI shall not freeze Excel.
- If the focus is on the GUI control, then it shall not be necessary to re-activate the Excel window before selecting the range

We will narrow to WinForms for the moment, in order to make it simple.

Facts
--------
- Open the form with the "Show" method from the Excel thread will make Excel freeze
- Open the form with the "Show" method from a separate thread will have the following consequence : when user starts typing in a box from the GUI, Excel steals the focus and user ends editing an Excel cell.
- Open the form with the "ShowDialog" method from the Excel thread will have the following consequence : if user focuses on a box in the GUI, then user must first re-activate Excel before selecting a range. So in the end, user must click twice in Excel in order to select the range.
- Open the form with the "Show" method from the Excel thread, passing the Excel handle as a parameter via a IWin32Window object works fine. Excel is not frozen when the GUI is alive, and user can select ranges into Excel without having to activate it first. However, if user starts editing a cell in Excel, or if Excel is busy for any reason, then the GUI is frozen because it shares the Excel thread. 

I noticed that the Bloomberg add-in contains a function helper GUI which behaves perfectly. See below some screenshots (I removed some graphical elements by precaution).

First, here is the color of the Excel title bar when the application is active:
![Active Excel](https://raw.github.com/Ron-Ldn/ExcelWizardLikeGUIs/master/Screenshots/Excel_active.png)

And here is the color of the Excel title bar when the application is not the active one (darker):
![Inactive Excel](https://raw.github.com/Ron-Ldn/ExcelWizardLikeGUIs/master/Screenshots/Excel_inactive.png)

The Bloomberg wizard is available from the add-in ribbon and looks like this:
![Inactive Excel](https://raw.github.com/Ron-Ldn/ExcelWizardLikeGUIs/master/Screenshots/func_wiz.png)

User can focus on one of the parameter boxes and then select a range into Excel, without having to activate Excel. Note that the title bar color shows that Excel is still inactive:
![Select Range in Excel](https://raw.github.com/Ron-Ldn/ExcelWizardLikeGUIs/master/Screenshots/Excel_selection.png)

Once the selection is made, the GUI will restore the focus to the selected box, and will populate the range address into it.

If user edits the formula menu in Excel, then Excel becomes active (look at title bar):
![Edit formula in Excel](https://raw.github.com/Ron-Ldn/ExcelWizardLikeGUIs/master/Screenshots/editing_formula_bar.png)

Note that when Excel is busy, for instance if user is editing the formula bar, then the GUI does not freeze. 


