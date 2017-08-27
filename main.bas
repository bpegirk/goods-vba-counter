Attribute VB_Name = "Module1"
Const DAY_GOOD_START_ROW_NUMBER = 2
Const DAY_GOOD_NAME_COLUMN_NUMBER = 2
Const DAY_GOOD_COUNT_COLUMN_NUMBER = 4
Const DAY_GOOD_AMOUNT_COLUMN_NUMBER = 5

Const GOOD_PAY_NAME_COLUMN_NUMBER = 1
Const GOOD_PAY_COUNT_COLUMN_NUMBER = 2
Const GOOD_PAY_AMOUNT_COLUMN_NUMBER = 3
Const sheetName = "ИТОГ ПРОДАЖ"
Sub makeGoodsPayCountSheet()
 
    createSheet (sheetName)
    Set sheet = Sheets(sheetName)
    
    ' clear cheet
    sheet.UsedRange.ClearContents
    
    sheet.Cells(1, 1).Value = "Название"
    sheet.Cells(1, 2).Value = "Кол-во"
    ' need add Microsoft Scripting Runtime library. (Add a reference to your project from the Tools...References menu in the VBE.)
    Dim d As Dictionary
    Set d = New Dictionary
    
    rowNum = 1
    For Each s In Sheets
        If (isDaySheet(s)) Then
            sheetRowNum = 2
                
            While (isGoodName(s.Cells(sheetRowNum, DAY_GOOD_NAME_COLUMN_NUMBER).Value))
                goodName = s.Cells(sheetRowNum, DAY_GOOD_NAME_COLUMN_NUMBER).Value
                goodCount = s.Cells(sheetRowNum, DAY_GOOD_COUNT_COLUMN_NUMBER).Value
                goodAmount = s.Cells(sheetRowNum, DAY_GOOD_AMOUNT_COLUMN_NUMBER).Value
                
                
                goodRow = d(goodName)
                If (goodRow = Empty) Then
                  goodRow = addGoodRow(sheet, goodName, rowNum)
                  d(goodName) = goodRow
                End If
                                
                sheet.Cells(goodRow, GOOD_PAY_COUNT_COLUMN_NUMBER).Value = (sheet.Cells(goodRow, GOOD_PAY_COUNT_COLUMN_NUMBER).Value + goodCount)
                sheet.Cells(goodRow, GOOD_PAY_AMOUNT_COLUMN_NUMBER).Value = (sheet.Cells(goodRow, GOOD_PAY_AMOUNT_COLUMN_NUMBER).Value + goodAmount)
                
                rowNum = WorksheetFunction.Max(rowNum, goodRow)
                sheetRowNum = sheetRowNum + 1
            Wend
        End If
    Next
End Sub

 Function isDaySheet(sheet)
    If (sheet.name = "Товары" Or sheet.name = "Шаблон" Or sheet.name = "ИТОГИ" Or sheet.name = "ЛИФОР" Or sheet.name = sheetName) Then
        isDaySheet = False
    Else
        isDaySheet = True
    End If
 End Function

                     
Sub createSheet(sheetName)
    finded = False
    For Each s In Sheets
        If (s.name = sheetName) Then
            finded = True
            Exit For
    End If
    Next
    If (finded = False) Then
        ActiveWorkbook.Sheets.Add(, Sheets(Sheets.Count)).name = sheetName
    End If
End Sub
   
Function isGoodName(name)
    If (name <> "" And name <> "ИТОГО") Then
        isGoodName = True
    Else
      isGoodName = False
    End If
End Function
  
          
Function addGoodRow(sheet, goodName, maxRowNumber)
  goodRowNumber = maxRowNumber + 1
  sheet.Cells(goodRowNumber, GOOD_PAY_NAME_COLUMN_NUMBER).Value = goodName
  sheet.Cells(goodRowNumber, GOOD_PAY_COUNT_COLUMN_NUMBER).Value = 0
  sheet.Cells(goodRowNumber, GOOD_PAY_AMOUNT_COLUMN_NUMBER).Value = 0
  addGoodRow = goodRowNumber
End Function


