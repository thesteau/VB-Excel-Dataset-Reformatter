Sub Main()
    Identification_Macro
End Sub

Sub Identification_Macro()
'
' Assumption is a raw save of the Essbase download - DO NOT RENAME THE SHEET NAMES - Keep it as is
' I know there is a lot of repetitive code - note that this was intentionally used to debug - feel free to 
' make a copy and cleanup

    Sheet_dup_and_mod

    Total = WorksheetFunction.Sum(ActiveSheet.Range("e:e"))

    lastrow = Range("A" & Rows.Count).End(xlUp).Row
    identification_of_row_difference(lastrow)
    Formulaics(lastrow)
    Spacer(lastrow)

    newlastrow = Range("A" & Rows.Count).End(xlUp).Row
    Sum_reapplication (newlastrow)
    Cleanup
    Intermediary_Final_Block newlastrow, Total
    Final_Block newlastrow
    
    M_Col_highlight newlastrow

End Sub

Sub Sheet_dup_and_mod()
    ActiveSheet.Select
    ActiveSheet.Copy Before:=Sheets(1)
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "raw"
    Sheets("Sheet1 (2)").Select
    Sheets("Sheet1 (2)").Name = "forecast"
End Sub

Sub identification_of_row_difference(lastrow)
'  Clear the first row, then apply formula to the fixed destination accordingly.
'  Paste this as values after
    Rows("1:1").Select
    Selection.ClearContents
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-5]<>R[1]C[-5],1,0)"
    Range("F3").Select
    Selection.AutoFill Destination:=Range("F3:F" & lastrow)
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("f3").Select
End Sub

Sub Formulaics(lastrow)
    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("F2").Select
    ActiveCell.FormulaR1C1 = "Comps & Benefits"
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With

    Range("G2").Select
    ActiveCell.FormulaR1C1 = "G&A Expenses"
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With

    Range("H2").Select
    ActiveCell.FormulaR1C1 = "Total"
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With

    Columns("C:C").Select
    With Selection
        .NumberFormat = "General"
        .Value = .Value
    End With
    Range("F3").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]<6061,RC[-1],0)"
    Range("G3").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]=0,RC[-2],0)"
    Range("H3").Select
    ActiveCell.FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"
    Range("I3").select
    ActiveCell.FormulaR1C1 = "=RC[-4]-RC[-1]"
    Range("F3:i3").Select
    Selection.AutoFill Destination:=Range("F3:i" & lastrow)
    Columns("F:G").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Columns("I:I").Select
    Application.CutCopyMode = False
    Selection.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
End Sub

Sub Spacer(lastrow)
For Each cell In ActiveSheet.Range("j1:j" & lastrow)
    If cell.Value = 1 Then
        cell.offset(1).Select
        Selection.EntireRow.Select
        Selection.Insert shift:=xlUp
        Selection.Insert shift:=xlUp
    End If
Next cell
End Sub

Sub Sum_reapplication(newlastrow)
    summer (newlastrow)
End Sub

sub summer(newlastrow)
For Each cell In ActiveSheet.Range("j1:j" & newlastrow)
    If cell.Value = 1 Then
        cell.Select
        top_select = Range(Selection, Selection.End(xlUp)).Count
        cell.Offset(1, -5).Select
        Selection.FormulaR1C1 = "=SUBTOTAL(9,R[-" & top_select & "]C:R[-1]C)"
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

        cell.Offset(1, -4).Select
        Selection.FormulaR1C1 = "=SUBTOTAL(9,R[-" & top_select & "]C:R[-1]C)"
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

        cell.Offset(1, -3).Select
        Selection.FormulaR1C1 = "=SUBTOTAL(9,R[-" & top_select & "]C:R[-1]C)"
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

        cell.Offset(1, -2).Select
        Selection.FormulaR1C1 = "=SUBTOTAL(9,R[-" & top_select & "]C:R[-1]C)"
    Selection.Font.Bold = True
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -4144960
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    End If
    top_select = 0
Next cell
End Sub

Sub Cleanup()
    Columns("A:D").EntireColumn.AutoFit
    Columns("E:H").Select
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Columns("E:H").EntireColumn.AutoFit

    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.ColumnWidth = 1.00

    Columns("H:H").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.ColumnWidth = 1.00

    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.ColumnWidth = 1.00

    Columns("F:F").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.ColumnWidth = 1.00

    Columns("N:N").clearcontents
End Sub

Sub Intermediary_Final_Block(newlastrow, Total)
    Range("A" & newlastrow).Offset(3, 1).Select
    Selection.Font.Bold = True
    Selection.value = "TOTAL: Comp & Benefits and G&A"

    Selection.offset(,3).select
    Selection.Font.Bold = True
    Selection.FormulaR1C1 = "=SUBTOTAL(9,R[-" & newlastrow & "]C:R[-1]C)"
    selection.entirecolumn.autofit

    Selection.offset(,2).select
    Selection.Font.Bold = True
    Selection.FormulaR1C1 = "=SUBTOTAL(9,R[-" & newlastrow & "]C:R[-1]C)"
    selection.entirecolumn.autofit

    Selection.offset(,2).select
    Selection.Font.Bold = True
    Selection.FormulaR1C1 = "=SUBTOTAL(9,R[-" & newlastrow & "]C:R[-1]C)"
    selection.entirecolumn.autofit

    Selection.offset(,2).select
    Selection.Font.Bold = True
    Selection.FormulaR1C1 = "=SUBTOTAL(9,R[-" & newlastrow & "]C:R[-1]C)"
    selection.entirecolumn.autofit

    Selection.offset(,2).select
    Selection.FormulaR1C1 = "=RC[-8]-RC[-2]"

    Range("A" & newlastrow).Offset(5, 1).Select
    Selection.value = "From original data dump"
    Selection.offset(,3).select
    Selection.value = Total

    Selection.offset(,8).select
    Selection.FormulaR1C1 = "=R[-2]C[-8]-RC[-8]"

End Sub

Sub Final_Block(newlastrow)
    Range("A" & newlastrow).Offset(7, 2).Select
    Selection.value = "Comp & Benefits:"

    Selection.offset(1,1).Select
    Selection.value = "Salaries & Wages"
    Selection.offset(,1).Select
    Selection.FormulaR1C1 = "=SUMIFS(R3C5:R"& newlastrow &"C5,R3C4:R"& newlastrow &"C4,RC[-1])"

    Selection.offset(1,-1).Select
    Selection.value = "Employee Bonuses"
    Selection.offset(,1).Select
    Selection.FormulaR1C1 = "=SUMIFS(R3C5:R"& newlastrow &"C5,R3C4:R"& newlastrow &"C4,RC[-1])"

    Selection.offset(1,-1).Select
    Selection.value = "Employer 401K Match"
    Selection.offset(,1).Select
    Selection.FormulaR1C1 = "=SUMIFS(R3C5:R"& newlastrow &"C5,R3C4:R"& newlastrow &"C4,RC[-1])"

    Selection.offset(1,-1).Select
    Selection.value = "Health Insurance"
    Selection.offset(,1).Select
    Selection.FormulaR1C1 = "=SUMIFS(R3C5:R"& newlastrow &"C5,R3C4:R"& newlastrow &"C4,RC[-1])"

    Selection.offset(1,-1).Select
    Selection.value = "Payroll Taxes"
    Selection.offset(,1).Select
    Selection.FormulaR1C1 = "=SUMIFS(R3C5:R"& newlastrow &"C5,R3C4:R"& newlastrow &"C4,RC[-1])"

    Selection.offset(1,-1).Select
    Selection.value = "Worker's Comp Insurance"
    Selection.offset(,1).Select
    Selection.FormulaR1C1 = "=SUMIFS(R3C5:R"& newlastrow &"C5,R3C4:R"& newlastrow &"C4,RC[-1])"

    Selection.offset(1,-1).Select
    Selection.value = "Total; Comp & Benefits"
    Selection.font.Bold = True
    Selection.offset(,1).Select
    Selection.FormulaR1C1 = "=SUM(R[-6]C:R[-1]C)"
    Selection.font.Bold = True

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    Selection.offset(,8).Select
    Selection.FormulaR1C1 = "=R[-11]C[-6]-RC[-8]"

End Sub

sub M_Col_highlight(newlastrow)
newlastrow = newlastrow + 14
    For Each cell In ActiveSheet.Range("m1:m" & newlastrow)
       If cell.Value = 0 and cell.HasFormula = TRUE Then
       cell.select
           With Selection.Interior
                  .Pattern = xlSolid
                  .PatternColorIndex = xlAutomatic
                  .Color = 65535
                  .TintAndShade = 0
                  .PatternTintAndShade = 0
           end With
       end if   
    next cell
End Sub
