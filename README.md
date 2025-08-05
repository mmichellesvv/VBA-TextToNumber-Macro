# VBA-TextToNumber-Macro
A simple VBA macro that highlights and converts numbers stored as text in Excel to actual numeric values, preserving the number format.

## Capabilities:
- highlight cells with numbers stored as text
- confirm and create a backup
- convert values using CDbl() while preserving decimal formatting
- works across all sheets in a workbook

## How to Use:

1. Open Excel.
2. Activate Developer mode following `File > Options > Costumise Ribbon, then check the box Developer in the right-hand list. Click `OK`.
2. Find Developer tab and click `Visual Basic` to open the VBA editor.
3. In the VBA editor, go to `File > Import File...`, and select the `TextToNumber.bas` file (which should be located in the same directory).
4. Go to `Developer > Macros`, select `TextNumberConverterWizard`, and click `Run`. (Alternatively, press Alt + F8, choose the macro, and click Run.)

You will be prompted to:
- Highlight problematic cells
- Optionally create a backup
- Convert text-formatted numbers to proper numeric format
