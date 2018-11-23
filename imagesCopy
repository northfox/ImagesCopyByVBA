Sub imagesCopy()
'
' imagesCopy Macro
'

  Dim originalSheet, targetSheet As Worksheet
  Dim imageName, originalSheetName, targetSheetName, originalBookName As String
  Dim i As Integer
 
  originalBookName = ActiveWorkbook.Name
  originalSheetName = ActiveSheet.Name
  Set originalSheet = Workbooks(originalBookName).Worksheets(originalSheetName)

  For Each targetSheet In ActiveWorkbook.Sheets
    If originalSheet.Name <> targetSheet.Name Then
      Dim Top_y, Left_x, Height_y, Width_x As Double
      originalSheet.Activate
      For i = 1 To ActiveSheet.Shapes.Count
        imageName = ActiveSheet.Shapes(i).Name

        Top_y = ActiveSheet.Shapes(imageName).Top
        Left_x = ActiveSheet.Shapes(imageName).Left
        Height_y = ActiveSheet.Shapes(imageName).Height
        Width_x = ActiveSheet.Shapes(imageName).Width

        originalSheet.Shapes(imageName).Copy
        targetSheet.Paste

        With targetSheet.Shapes(1)
          .Left = Left_x
          .Top = Top_y
          .Height = Height_y
          .Width = Width_x
        End With
      Next i
    End If
  Next
  MsgBox " 画像を全シートにコピーしました (シート数 : " & ActiveWorkbook.Sheets.Count - 1 & ") 。 "
End Sub
