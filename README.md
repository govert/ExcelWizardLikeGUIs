# ExcelWizardLikeGUIs
Research on how to build GUIs which can select ranges into Excel, within an Excel add-in

This project is for research only, it is not meant to be an usable tool in the end.
The purpose is to find out how to make a GUI available from an Excel add-in, which will allow to select a range in order to populate the address into some control within that GUI.

- The GUI shall not freeze Excel.
- If the focus is on the GUI control, then it shall not be necessary to re-activate the Excel window before selecting the range

We will narrow to WinForms for the moment, in order to make it simple.

Some facts :
1) Open the form with the "Show" method from the Excel thread => this freezes Excel
2) Open the form with the "Show" method from a separate thread => when user starts typing in a box from the GUI, Excel steals the focus and user ends editing an Excel cell.
3) Open the form with the "ShowDialog" method from the Excel thread => if user focuses on a box in the GUI, then user must first re-activate Excel before selecting a range. So in the end, user must click twice in Excel in order to select the range.
4) Open the form with the "Show" method from the Excel thread, passing the Excel handle as a parameter via a IWin32Window object. This works fine, Excel is not frozen when the GUI is alive, and user can select ranges into Excel without having to activate it first. However, if user start editing a cell in Excel, or if Excel is busy for any reason, then the GUI is frozen because it shares the Excel thread. 
