# MaxDrawdown
MS Excel VBA Function to calculate the maximum drawdown of a table of asset prices

Installation Instructions:
* Locate or download the module file MaxDrawdown.bas
* Open the Excel workbook where you want to import the MaxDrawdown.bas module
* Enable the Developer Tab:
    - Go to File > Options.
    - In the Excel Options dialog box, select Customize Ribbon.
    - Check the box for Developer in the right pane and click OK.
* Open the VBA Editor:
    - Click on the Developer tab.
    - Select Visual Basic or press Alt + F11 to open the Visual Basic for Applications (VBA) editor.
* Import .bas File:
    - In the VBA editor, go to File > Import File.
    - Navigate to the location of your MaxDrawdown.bas file, select it, and click Open.
* Using the MaxDrawdown Function:
    - The imported module will appear under Modules in the VBA Project Explorer.
    - Close the VBA editor and return to your Excel workbook.
    - To use the MaxDrawdown function, enter the following formula in a cell:
    
                        =MaxDrawdown(A1:D10)
    - Replace A1:D10 with the actual range of your asset prices, where the first column contains the dates and the first row contains the asset names.
* Interpreting the Results:
    - The function will return a range with one column for each asset, containing:
        + the maximum drawdowns, 
        + the peak date prior to the maximum drawdown, 
        + the trough date, and 
        + the recovery date.