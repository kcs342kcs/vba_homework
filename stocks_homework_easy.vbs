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
  call stocks_easy
end sub

sub Scripts():
  call AddScriptingLibrary
end sub

Sub stocks_easy():

 Dim lastRow As Long
 Dim lastColumn As Long
 dim ws as worksheet
 ' remember starting worksheet
 dim starting_ws as worksheet
 set starting_ws = ActiveSheet
 Dim rowC As Long
 Dim i As Long
 ' define an associative array (scripting.dictionary) to store ticker => volume info
 Dim stockCounter As Scripting.Dictionary
 Set stockCounter = New Scripting.Dictionary
 stockCounter.CompareMode = vbBinaryCompare
 
 for each ws in ThisWorkbook.worksheets
  ws.Activate
  'empty dict for this worksheet
  stockCounter.RemoveAll
 ' get the total row and column count for loop interation
  lastRow = getRC("r", ActiveSheet)
  lastColumn = getRC("c", ActiveSheet)

 ' put new headers in first row for new info
  Cells(1, lastColumn + 2).Value = "Ticker"
  Cells(1, lastColumn + 3).Value = "Total Stock Volume"

  ' auto fit new columns
  Columns(lastColumn + 2).AutoFit 
  Columns(lastColumn + 3).AutoFit 

 ' instantiate scripting.dictionary with ticker/volume info
  For rowC = 2 To lastRow
    stockCounter(Cells(rowC, 1).Value) = stockCounter(Cells(rowC, 1).Value) + Cells(rowC, 7).Value
  Next rowC
 
 ' put ticker/volume info onto worksheet
  For i = 0 To stockCounter.Count - 1
    Cells(i + 2, lastColumn + 2).Value = stockCounter.Keys(i)
    Cells(i + 2, lastColumn + 3).Value = stockCounter.Items(i)
  Next i
 Next
 'reactivate original worksheet
 starting_ws.Activate
 
End Sub

