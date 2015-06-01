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
The following scenarios have been tested with Excel 2010 and Excel 2013, using the .Net framework 3.5 and 4.0. The behaviour does not change depending on the version of Excel or of the .Net framework.

- Scenario 1/ form.Show() is called from the Excel thread. In that case, Excel is not frozen when the GUI is running. It is possible to select a range in Excel without having to activate Excel first. If Excel is busy, for instance if a cell or the formula bar is in edit mode, then the GUI is frozen.
- Scenario 2/ form.ShowDialog() from the Excel thread. In that case, Excel is frozen when the GUI is running. It is not possible to select a range into Excel.
- Scenario 3/ form.Show() is called from a separate thread. The GUI disappears when the thread exits.
- Scenario 4/ form.ShowDialog() is called from a separate thread. In that case, Excel is not frozen when the GUI is running. If Excel is busy, for instance if a cell or the formula bar is in edit mode, then the GUI is not frozen and can be used with no issue. In order to select a range into Excel, user must first activate Excel. 
- Scenario 5/ form.Show(IWin32Window) is called from the Excel thread, passing the Hwnd value of the Excel application. Same behaviour as in 1.
- Scenario 6/ form.ShowDialog(IWin32Window) from the Excel thread., passing the Hwnd value of the Excel application. Same behaviour as in 2.
- Scenario 7/ form.Show(IWin32Window) is called from a separate thread, passing the Hwnd value of the Excel application. This is bad. The GUI disappears as soon as the thread exits, but Excel remains frozen.
- Scenario 8/ form.ShowDialog(IWin32Window) is called from a separate thread, passing the Hwnd value of the Excel application. It is not possible to activate Excel ; the GUI re-activates itself automatically (with a little blink).


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


