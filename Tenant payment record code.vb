 Option Explicit
Dim MCells As Variant
Dim JCells As Variant
Dim RCells As Variant
Dim OCells As Variant
Dim FCells As Variant
Dim BCells As Variant
Dim CCells As Variant
Dim NCells As Variant

Sub Reset_Tenant_Payement()
Dim Response As Variant

Response = MsgBox("Do you want to reset?", vbYesNo + vbCritical)
 
 If Response = vbYes Then


   '==============================
   ' Clear the sheet and contents
   '=============================
     
   ActiveSheet.Range("A2:H1020").Select
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    
    ActiveSheet.Range("F2:H1020").Select
    With Selection
        .ClearContents
    End With
    
    '==============================================
    ' RESET YELLOW CELLS
    '==============================================
    
    Set MCells = ActiveSheet.Range("A1:H1020").Find("MAYFAIR ESTATE", LookIn:=xlValues)
    
    MCells.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    Set JCells = ActiveSheet.Range("A1:H1020").Find("JS MANAGEMENT", LookIn:=xlValues)
    JCells.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
   
    Set RCells = ActiveSheet.Range("A1:H1020").Find("RAJ", LookIn:=xlValues)
    RCells.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    '=============================================================
    ' RESET INDICATION CELLS
    '==============================================================
    Set FCells = ActiveSheet.Range("A1:H1020").Find("FULL PAID", LookIn:=xlValues)
    FCells.Offset(, -1).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 6750156
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
    Set BCells = ActiveSheet.Range("A1:H1020").Find("BALANCE IS DUE", LookIn:=xlValues)
     BCells.Offset(, -1).Select
   With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
    Set CCells = ActiveSheet.Range("A1:H1020").Find("PAID CASH", LookIn:=xlValues)
    CCells.Offset(, -1).Select
     With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16751052
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Set OCells = ActiveSheet.Range("A1:H1020").Find("MOVING OUT", LookIn:=xlValues)
    OCells.Offset(, -1).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16764159
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Set NCells = ActiveSheet.Range("A1:H1020").Find("NEW TENANT", LookIn:=xlValues)
    NCells.Offset(, -1).Select
     With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16777062
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    '======================================
    ' CLEAR CHECK BOXES
    '=====================================
    ActiveSheet.CheckBoxes.Value = False
    
Else
    Exit Sub
End If
    
End Sub

Sub MoveOut_Code()

ActiveSheet.Range(ActiveCell, ActiveCell.Offset(, 7)).Select

 With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 16764159
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub BlueLine_Code()

   Range(ActiveCell, ActiveCell.Offset(, 7)).Select
   
   With Selection.Font
       .Bold = True
       .Italic = True
       .Size = 12
       .Name = "Calibri"
   End With
   
   With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
   With Selection
     .HorizontalAlignment = xlCenter
     .VerticalAlignment = xlBottom
   End With
   
   With Selection
     .RowHeight = 24
   End With
   
   ActiveCell.Value = "Due"
   ActiveCell.Offset(, 1).Value = "Address"
  'ActiveCell.Offset(, 2).Value = "Landlord"
   ActiveCell.Offset(, 2).Value = "Name"
   ActiveCell.Offset(, 3).Value = "Rent"
   ActiveCell.Offset(, 4).Value = "Payment"
   ActiveCell.Offset(, 5).Value = "Date"
   ActiveCell.Offset(, 6).Value = "Amt.Paid"
   ActiveCell.Offset(, 7).Value = "Recd."
   
End Sub
Sub NewChanges_Code()

ActiveSheet.Range(ActiveCell, ActiveCell.Offset(, 7)).Select

 With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        '.Color = 16764190
        .Color = 16776960
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub Header()
 ActiveWindow.View = xlPageLayoutView
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    Selection.Font.Bold = True
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = "&""-,Bold""TENANT'S PAYMENT-   23"
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = "&""-,Bold""TENANT'S PAYMENT-   23"
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    
End Sub

Sub RentPaid_Code()
   
   Dim ReturnValue As Boolean
   
   ReturnValue = CheckCell(ActiveCell.Value)
   
   If ReturnValue = False Then
      MsgBox ("Not Paid - Exiting sub")
      Exit Sub
   Else
      MsgBox ("Rent Paid")
     
      
      If Application.WorksheetFunction.IsText(ActiveCell) = True Then
      
           MsgBox ("Full rent has been paid.")
           
           Range(ActiveCell, ActiveCell.Offset(, -6)).Select
           
           With Selection.Interior
          .Pattern = xlSolid
          .PatternColorIndex = xlAutomatic
          .Color = 6750156
          .TintAndShade = 0
          .PatternTintAndShade = 0
          End With
         
         
         ActiveCell.Offset(, 7).Select
         With Selection.Interior
          .Pattern = xlSolid
          .PatternColorIndex = xlAutomatic
          .Color = 6750156
          .TintAndShade = 0
          .PatternTintAndShade = 0
         End With
      Else
          
        MsgBox ("Balance is due.")
          
        Range(ActiveCell, ActiveCell.Offset(, -6)).Select
        
        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 39423
        .TintAndShade = 0
        .PatternTintAndShade = 0
        End With
        
        ActiveCell.Offset(, 7).Select
        With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 39423
        .TintAndShade = 0
        .PatternTintAndShade = 0
        End With
        
        
      End If
    End If
End Sub

Private Function CheckCell(CellValue) As Boolean

  If IsEmpty(CellValue) Then
  
     CheckCell = False
  Else
     CheckCell = True
  End If


End Function
