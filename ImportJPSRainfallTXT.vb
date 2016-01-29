Sub ImportJPSRainfallTXT()

Dim ThisWB As Workbook
Dim TemplateWB As Workbook
Dim NewDataWS As Worksheet
Dim TemplateWS As Worksheet
Dim RFWS As Worksheet

Set TemplateWB = ThisWorkbook


''''new workbook
Dim InitialName As String
Dim sFileSaveName As Variant

Set ThisWB = Workbooks.Add



''define Sheet1 in new workbook
Set TemplateWS = TemplateWB.Sheets("RFData")
Set RFWS = ThisWB.Sheets(1)

With RFWS
    .Name = "RFData"
    .Range("A1") = "Stn_No"
    .Range("B1") = "Date"
    .Range("C1") = "Depth (mm)"
End With



''''import txt file

On Error GoTo ErrorHandler

'''dialog box to choose folder

  'Declare variables
  Dim strOriginalPath As String
  Dim x As Variant
  Dim strSelectedPath As String
  
  
  'Remember original path
  strOriginalPath = CurDir
    
  'Show dialog and let user browse to a directory
  x = Application.GetOpenFilename(Title:="Select (any) txt file in the folder and click Open")
   
  'Get current directory
  strSelectedPath = CurDir

    'specify directory
    Dim FolderName As String, FName As String
    FolderName = strSelectedPath & "\"
    
    ''loop through all *.txt files in folders
    FName = Dir(FolderName & "*.txt") 'gets the list of *.txt files
            
                While FName <> ""
                                                            
                    'import txt
                    Set NewDataWS = ThisWB.Sheets.Add(After:=ThisWB.Sheets(ThisWB.Sheets.Count))
                    With NewDataWS.QueryTables.Add(Connection:="TEXT;" & FolderName & FName _
                        , Destination:=Range("$A$1"))
                        .FieldNames = True
                        .RowNumbers = False
                        .FillAdjacentFormulas = False
                        .PreserveFormatting = True
                        .RefreshOnFileOpen = False
                        .RefreshStyle = xlInsertDeleteCells
                        .SavePassword = False
                        .SaveData = True
                        .AdjustColumnWidth = True
                        .RefreshPeriod = 0
                        .TextFilePromptOnRefresh = False
                        .TextFilePlatform = 437
                        .TextFileStartRow = 1
                        .TextFileParseType = xlFixedWidth
                        .TextFileTextQualifier = xlTextQualifierDoubleQuote
                        .TextFileConsecutiveDelimiter = False
                        .TextFileTabDelimiter = True
                        .TextFileSemicolonDelimiter = False
                        .TextFileCommaDelimiter = False
                        .TextFileSpaceDelimiter = False
                        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
                        .TextFileFixedColumnWidths = Array(9, 4, 6, 7, 8, 8, 8, 8, 8, 9, 7, 8, 8, 10)
                        .TextFileTrailingMinusNumbers = True
                        .Refresh BackgroundQuery:=False
                    End With
                   
                    
                    ''copy data from imported sheet
                    Dim RFDataPreLastRow As Long
                    Dim RFDataLastRow As Long
                    Dim i As Long
                    Dim Month As Long
                    Dim j As Long
                    Dim StartYr As Variant
                    Dim AfterYrs As Long
                    Dim StartRow As Long
                    Dim YrFirstRow As Long
                    Dim YrEndRow As Long
                    Dim ThisYr As Integer

                    RFDataPreLastRow = RFWS.Cells(Rows.Count, "A").End(xlUp).Row    'last row before copy 1 yr data
                    DataRow = NewDataWS.Cells(Rows.Count, "C").End(xlUp).Row
                    StartRow = NewDataWS.Range("A1:A100").Find("Rain").Row + 4   'search 'Rain' in cell,find data start row
                    j = DataRow / 43    ' j = total rows/43 rows = total years in dataset
                                        
                    NewDataWS.Name = Mid(FName, 1, 7)     'rename worksheet
                    ThisYr = Year(Now())
                    
                    
                    For i = 1 To j
                    
                    
                        'determine first row of data for each year
                        If i = 1 Then
                            YrFirstRow = StartRow * i + 33 * (i - 1)
                        Else
                            YrEndRow = YrFirstRow + 33
                            YrFirstRow = NewDataWS.Range("A" & YrEndRow & ":A" & (YrEndRow + 20)).Find("Rain").Row + 4
                        End If
                        
                        
                        'determine year
                        If NewDataWS.Range("E" & (YrFirstRow - 5)).Text < (ThisYr - 2000) Then
                            StartYr = NewDataWS.Range("E" & (YrFirstRow - 5)).Text + 2000
                        Else
                            StartYr = NewDataWS.Range("D" & (YrFirstRow - 5)).Characters(6, 2).Text & _
                                NewDataWS.Range("E" & (YrFirstRow - 5)).Text
                        End If
                        
                        
                        'copy data each month - different number of rows each month
                        For Month = 1 To 12
                                                                                        
                        RFDataLastRow = RFWS.Cells(Rows.Count, "C").End(xlUp).Row
                        NewDataWS.Activate
                        
                            If Month = 2 Then
                                If StartYr Mod 4 = 0 Then
                                    NewDataWS.Range(Cells(YrFirstRow, Month + 2), _
                                        Cells(YrFirstRow + 28, Month + 2)).Copy
                                        
                                Else
                                    NewDataWS.Range(Cells(YrFirstRow, Month + 2), _
                                        Cells(YrFirstRow + 27, Month + 2)).Copy
                                End If
                            ElseIf Month = 4 Or Month = 6 Or Month = 9 Or Month = 11 Then
                                NewDataWS.Range(Cells(YrFirstRow, Month + 2), _
                                    Cells(YrFirstRow + 29, Month + 2)).Copy
                            
                            Else
                                NewDataWS.Range(Cells(YrFirstRow, Month + 2), _
                                    Cells(YrFirstRow + 30, Month + 2)).Copy
                            
                            End If
                            
                            RFWS.Range("C" & RFDataLastRow + 1).PasteSpecial Paste:=xlPasteValues   'paste data
        
                        Next Month

                    
                    ''fill station date
                    RFDataPreLastRow = RFWS.Cells(Rows.Count, "B").End(xlUp).Row
                    RFDataLastRow = RFWS.Cells(Rows.Count, "C").End(xlUp).Row
                    With RFWS
                        .Range("B" & (RFDataPreLastRow + 1)).FormulaR1C1 = "1/1/" & StartYr       'date data starts
                        .Range("B" & (RFDataPreLastRow + 1)).AutoFill _
                            Destination:=.Range("B" & (RFDataPreLastRow + 1) & ":B" & RFDataLastRow), _
                            Type:=xlFillDefault                                                 'fill rest of date
                    End With
                    
                    
                    Next i
                    
                    
                    ''fill station number
                    RFDataPreLastRow = RFWS.Cells(Rows.Count, "A").End(xlUp).Row
                    RFDataLastRow = RFWS.Cells(Rows.Count, "C").End(xlUp).Row
                    
                    With RFWS
                        .Range("A" & RFDataPreLastRow + 1) = NewDataWS.Name                     'station number
                        .Range("A" & RFDataPreLastRow + 1).AutoFill _
                            Destination:=.Range("A" & RFDataPreLastRow + 1 & ":A" & RFDataLastRow)   'fill station number
                    End With
                                    
                
                
                FName = Dir()
                
                Wend    'Loop

        
''save workbook
RFWS.Columns("A:C").AutoFit
RFWS.Activate

InitialName = "JPS_RF_"

DoEvents    'allow all the background activity to complete before save

sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=ThisWorkbook.Path & "\" & InitialName, _
    fileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Save new time-series data as")

If sFileSaveName <> False Then
    ThisWB.SaveAs sFileSaveName
End If

ThisWB.Save
Application.CutCopyMode = False


ErrorHandler:
    If Err.Number <> 0 Then
      MsgBox "Error Number " & Err.Number & vbNewLine & Err.Description
    End If


End Sub
