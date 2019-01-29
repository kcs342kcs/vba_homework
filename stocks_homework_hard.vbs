 'gets the row or column number of the last row/column with data
 'only to be used on a well formatted data sheet
 'blank cells/columns in the middle of a row/column screw it up
 Function getRC(ByVal rowColumn As String, ByVal sht As Worksheet) As Long
 
  If (StrComp(rowColumn, "c", vbTextCompare)) Then
    getRC = sht.Range("A1").CurrentRegion.Rows.Count
  ElseIf (StrComp(rowColumn, "r", vbTextCompare)) Then
    getRC = sht.Range("A1").CurrentRegion.Columns.Count
  Else
    getRC = 1
  End If

 End Function
  
  ' turns on microsoft scripting runtime library, if necessary, to allow the use of scripting dictionarys
  ' is dependent upon having trust access to the vba project model enabled (see below)
  ' File -> options -> trust center -> trust center settings -> macro settings -> trust access to the VBA project model

 Private Function AddScriptingLibrary() As Boolean
    dim ref as object
    Const GUID As String = "{420B2830-E718-11CF-893D-00A0C9054228}"
    
    On Error GoTo errHandler
    for each ref in ActiveWorkbook.VBProject.References
       if (ref.name = "Scripting") then
        AddScriptingLibrary = True
        exit Function
       else
         'ThisWorkbook.VBProject.References.AddFromGuid GUID, 1, 0
         AddScriptingLibrary = False
         'Exit Function
       end If
    next ref
    if (AddScriptingLibrary = False) then
      ThisWorkbook.VBProject.References.AddFromGuid GUID, 1, 0
      exit Function
    end If
errHandler:
    MsgBox Err.Description
    
End Function

sub RunAll():
  call Scripts
  call stocks_hard
end sub

sub Scripts():
  call AddScriptingLibrary
end sub

Sub stocks_hard():
 Dim lastRow As Long
 Dim lastColumn As Long
 Dim rowC As Long
 Dim i As Long
 Dim openP As Double
 Dim closeP As Double
 Dim lastTicker As String
 Dim ws As Worksheet
 Dim starting_ws As Worksheet
 Set starting_ws = ActiveSheet
 Dim stockCounter As Scripting.Dictionary
 Set stockCounter = New Scripting.Dictionary
 stockCounter.CompareMode = vbBinaryCompare
 Dim openPrice As Scripting.Dictionary
 Dim closePrice As Scripting.Dictionary
 Set openPrice = New Scripting.Dictionary
 Set closePrice = New Scripting.Dictionary
 openPrice.CompareMode = vbBinaryCompare
 closePrice.CompareMode = vbBinaryCompare
 Dim priceDiff As Double
 Dim keyS As String
 Dim pricePercent As Double
 Dim maxVol as Double
 dim maxVolTicker as String
 dim lowPercent as Double
 dim lowPercentTicker as String
 dim highPercent as Double
 dim highPercentTicker as String
 
 ' garbage values for first run through loop
 closeP = 0
 lastTicker = "garbage"

 For Each ws In ThisWorkbook.worksheets
  ws.Activate

  ' clear varibles for new worksheet
  stockCounter.RemoveAll
  openPrice.RemoveAll
  closePrice.RemoveAll
  maxVol = 0
  maxVolTicker = ""
  lowPercent = 0
  lowPercentTicker = ""
  highPercent = 0
  highPercentTicker = ""
  
  ' get dimensions of data on sheet and write column headers

  lastRow = getRC("r", ActiveSheet)
  lastColumn = getRC("c", ActiveSheet)
  Cells(1, lastColumn + 2).Value = "Ticker"
  Cells(1, lastColumn + 3).Value = "Yearly Change"
  Cells(1, lastColumn + 4).Value = "Percent Change"
  Cells(1, lastColumn + 5).Value = "Total Stock Volume"
  cells(1, lastColumn + 9).Value = "Ticker"
  cells(1, lastColumn + 10).value = "value"
  cells(2, lastColumn + 8).value = "Greatest % Increase"
  cells(3, lastColumn + 8).value = "Greatest % Decrease"
  cells(4, lastColumn + 8).value = "Greatest Total Volume"

  ' auto fit new columns
  Columns(lastColumn + 2).AutoFit
  Columns(lastColumn + 3).AutoFit
  Columns(lastColumn + 4).AutoFit
  Columns(lastColumn + 5).AutoFit
  Columns(lastColumn + 8).AutoFit
  Columns(lastColumn + 9).AutoFit
  Columns(lastColumn + 10).AutoFit

  ' run through and parse data

  For rowC = 2 To lastRow
   ' check to see if ticker value exists in dict 
   If (stockCounter.Exists(Cells(rowC, 1).Value)) Then
    stockCounter(Cells(rowC, 1).Value) = stockCounter(Cells(rowC, 1).Value) + Cells(rowC, 7).Value
    closeP = Cells(rowC, 6).Value
    lastTicker = Cells(rowC, 1).Value
   Else
    openP = Cells(rowC, 3).Value
    openPrice(Cells(rowC, 1).Value) = openP
    closePrice(lastTicker) = closeP
    stockCounter(Cells(rowC, 1).Value) = stockCounter(Cells(rowC, 1).Value) + Cells(rowC, 7).Value
   End If
  Next rowC
  
  ' add last close price to dict
  closePrice(lastTicker) = closeP
  
  ' run through parsed data, place on sheet and format
  For i = 0 To stockCounter.Count - 1
    keyS = stockCounter.keyS(i)
    priceDiff = closePrice(keyS) - openPrice(keyS)
    If (openPrice(keyS) = 0) Then
      pricePercent = 0
      priceDiff = 0
    Else
      pricePercent = (priceDiff / openPrice(keyS))
    End If
    Cells(i + 2, lastColumn + 2).Value = stockCounter.keyS(i)
    Cells(i + 2, lastColumn + 3).Value = priceDiff
    if (priceDiff >= 0) Then
      Cells(i + 2, lastColumn + 3).Interior.ColorIndex = 10
    Else
      Cells(i + 2, lastColumn + 3).Interior.ColorIndex = 3
    end if

    Cells(i + 2, lastColumn + 4)= format(pricePercent, "Percent")
    Cells(i + 2, lastColumn + 5).Value = stockCounter.Items(i)
    if (stockCounter.Items(i) > maxVol) Then
      maxVolTicker = stockCounter.keyS(i)
      maxVol = stockCounter.Items(i)
    end if
    if (pricePercent > highPercent) Then
      highPercent = pricePercent
      highPercentTicker = stockCounter.keyS(i)
    end If
    if (pricePercent < lowPercent) Then
      lowPercent = pricePercent
      lowPercentTicker = stockCounter.keyS(i)
    end if
  Next i
  'write sheet totals
  cells(2, lastColumn + 9).value = highPercentTicker
  cells(2, lastColumn + 10) = format(highPercent, "Percent")
  cells(3, lastColumn + 9).value = lowPercentTicker
  cells(3, lastColumn + 10) = format(lowPercent, "Percent")
  cells(4, lastColumn + 9).value = maxVolTicker
  cells(4, lastColumn + 10).value = maxVol

 Next
 'activate original worksheets
 starting_ws.Activate
 
End Sub
