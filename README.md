# Excel Wingdings Checkbox VBA Script

This repository contains a simple VBA script to toggle between Wingdings symbols (used as checkboxes) in Excel by double-clicking on cells. The symbols 'o' and 'x' are used as unchecked and checked boxes, respectively.

## Features
- Toggle between checkbox symbols (like `o` and `x`) in **Column C** by double-clicking.
- No need to enable the Developer tab in Excel.
- Fully customizable for tracking data or creating dynamic checklists.

## Instructions
1. Open your Excel workbook.
2. Press `Alt + F11` to open the VBA editor.
3. Add code and adjust the column range if needed (currently set to **Column C**).

### Code
```vba
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' Specify the range where the checkboxes are located (column C)
    If Not Intersect(Target, Me.Range("C:C")) Is Nothing Then
        Cancel = True ' Prevents entering edit mode on double-click
        
        ' Check if the font is Wingdings and if it is a checkbox (either "o" or "x")
        If Target.Font.Name = "Wingdings" Then
            If Target.Value = "o" Then
                ' If the value is "o", change it to "x" (checked box)
                Target.Value = "x"
            ElseIf Target.Value = "x" Then
                ' If the value is "x", change it back to "o" (unchecked box)
                Target.Value = "o"
            End If
        End If
    End If
End Sub
```

## Future Features
- Plans to extend the script to other Wingdings symbols in different columns.
- You can customize this further by adding features like user timestamps, row locking, and more!

## License
This project is licensed under the MIT License.
