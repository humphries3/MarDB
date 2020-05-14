Attribute VB_Name = "mod_app"
Option Explicit
'''''''''''''''

Public Const MOD_DATE = "2020-02-23"
Public Const MOD_VER = 3
''''''''''''''''''''''''
'V3 recognize Excel date in DATE COL
''''''''''''''''''''''''''''''''''''

'Noteworthy links:
''''''''''''''''''
'REGEX: https://stackoverflow.com/questions/47613786/vba-regex-split-with-variable
'(REGEX should be used for enhancements needing pattern matching)



Public Const BLANK = " "
Public Const PERIOD = "."
Public Const COMMA = ","
Public Const HYPHEN = "-"
Public Const SLASH = "/"
Public Const ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const NUMBERS = "0123456789"

Public Type typShCtl
    Sh As Worksheet
    RowLast As Long
    RowCurr As Long
    ColMap As New Scripting.Dictionary
    End Type
    
Public Type typDateStore
    Date1 As Variant
    Date2 As Variant
    DateRange As Variant
    End Type

'ENUM over date range types,
'also (number of data components - 1) required to describe each range:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const DateRange_Y = 0
Public Const DateRange_M = 1
Public Const DateRange_D = 2

'Indexes into DATESTORE array, by date source:
''''''''''''''''''''''''''''''''''''''''''''''
Public Const DateStore_DATE = 0
Public Const DateStore_YYMM = 1
Public Const DateStore_DESC = 2
Public Const DateStore_WHER = 3

Public Type typNewData
    AccessRow As Long
    LibNam As String
    LibSeq As Long
    Pg As Long
    Ph As Long
    Roll As Long
    City As String
    State As String
    Desc As String
    Notes As String
    DateStore(0 To 3) As typDateStore
    DateStoreBest As Variant
    End Type
    
Public shPhotosAcc As typShCtl
Public shPhotosLog As typShCtl
Public shPhotosCnv As typShCtl
Public wbUser As Workbook
Public DateChars As New Scripting.Dictionary
Public AlphaHash As New Scripting.Dictionary
Public NumerHash As New Scripting.Dictionary
Public currYr As Long

Sub Run()
'''''''''
Dim ColNo As Long
Dim BlankRows As Long
Dim IndexRows As Long
Dim ErrorRows As Long
Dim NewColNames
Dim LogColNames
Dim rngAccess As Range
Dim Row As Long
Dim Rng As Range
Dim oLibNam As String
Dim oLibSeq As Long
Dim newData As typNewData
Dim DateChar
Dim SortRng As Range
''''''''''''''''''''
currYr = Year(Now)

NewColNames = Array( _
    "Access", _
    "Library", _
    "Album", _
    "Pg", _
    "Ph", _
    "Roll", _
    "DR", _
    "DS", _
    "Date(Start)", _
    "Date(End)", _
    "City", _
    "State", _
    "Description", _
    "Notes")

LogColNames = Array( _
    "Time", _
    "Row (Access)", _
    "Message", _
    "Column Data")
    
DateChars.RemoveAll

'Clear Log and output sheets:
'''''''''''''''''''''''''''''
With shPhotosLog
    .Sh.Cells.Clear
    .Sh.Rows(1).Font.Bold = True
    .RowLast = 1
    .ColMap.RemoveAll
    For ColNo = 1 To UBound(LogColNames) + 1
        .Sh.Cells(.RowLast, ColNo) = LogColNames(ColNo - 1)
        .ColMap.Add LogColNames(ColNo - 1), ColNo
        Next ColNo
    .Sh.Columns(.ColMap("Row (Access)")).NumberFormat = "#"
    End With

With shPhotosCnv
    .Sh.Cells.Clear
    .Sh.Rows(1).Font.Bold = True
    .Sh.Cells.VerticalAlignment = xlBottom ' to line up with excel row numbers
    .RowLast = 1
    .ColMap.RemoveAll
    For ColNo = 1 To UBound(NewColNames) + 1
        .Sh.Cells(.RowLast, ColNo) = NewColNames(ColNo - 1)
        .ColMap.Add NewColNames(ColNo - 1), ColNo
        Next ColNo
    .Sh.Columns(.ColMap("roll")).NumberFormat = "#"
    Set Rng = .Sh.Cells(1, .ColMap("dr"))
    Rng.AddComment "Date Range: D=DAY M=Month Y=Year"
    Set Rng = .Sh.Cells(1, .ColMap("ds"))
    Rng.AddComment "Date Source: Original column(s) from which date was derived"
    End With
    

Call logMsg(0, "I-001 Conversion started", "")
Set rngAccess = shPhotosAcc.Sh.Cells(1, 1).CurrentRegion
Call logMsg(0, "I-002 " & shPhotosAcc.Sh.Name & ": " & rngAccess.Address, "")
shPhotosAcc.RowLast = rngAccess.Rows.Count

Application.ScreenUpdating = False

For shPhotosAcc.RowCurr = 2 To shPhotosAcc.RowLast

    'instantiate new row data with i/p seq:
    '''''''''''''''''''''''''''''''''''''''
    newData = initNewData(shPhotosAcc.RowCurr)
    
    Select Case True
    
        'weed out rows with INDEX in LIBRARY column or rows missing most data:
        Case Not parseCandidate(shPhotosAcc, BlankRows, IndexRows)
        
        'normalize library name and isolate album number in new ALBUM col:
        Case Not parseLibAlbum(shPhotosAcc, newData)
        
        'validate Pg Ph for numerics:
        '''''''''''''''''''''''''''''
        Case Not parsePgPh(shPhotosAcc, newData)
        
        'normalize roll number (blanks and S -> 0):
        '''''''''''''''''''''''''''''''''''''''''''
        Case Not parseRoll(shPhotosAcc, newData)
        
        'parse city, state and possible date from WHERE col:
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case Not parseWhere(shPhotosAcc, newData)
        
        'parse possible date from DESC col:
        '''''''''''''''''''''''''''''''''''
        Case Not parseDescNotes(shPhotosAcc, newData)
        
        'look for date data in MM, YY, and DATE columns, then choose best one:
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case Not ChooseDate(shPhotosAcc, newData)
        
        'got a complete row, output:
        ''''''''''''''''''''''''''''
        Case Else
        Call addNewData(shPhotosCnv, newData)
        End Select
        
    Next shPhotosAcc.RowCurr
    
Application.ScreenUpdating = True
    
With shPhotosLog
    .Sh.Cells.AutoFilter
    .Sh.Columns.AutoFit
    .Sh.Activate
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
        End With

    End With
    
With shPhotosCnv

    If frmCP.ctlSortData Then
        Set SortRng = .Sh.Cells(1, 1).CurrentRegion
        SortRng.Sort key1:=.Sh.Cells(1, 1), order1:=xlAscending, _
                     key2:=.Sh.Cells(1, 2), order2:=xlAscending, _
                     key3:=.Sh.Cells(1, 3), order3:=xlAscending, _
                     Header:=xlYes, MatchCase:=False
        Call logMsg(0, "I-003 Sort range is " & SortRng.Address, "")
        End If

    .Sh.Cells.AutoFilter
    .Sh.Columns(.ColMap("date(start)")).NumberFormat = "mm/dd/yy"
    .Sh.Columns(.ColMap("date(start)")).HorizontalAlignment = xlCenter
    .Sh.Columns(.ColMap("date(end)")).NumberFormat = "mm/dd/yy"
    .Sh.Columns(.ColMap("date(end)")).HorizontalAlignment = xlCenter
    .Sh.Columns.AutoFit
    .Sh.Columns(.ColMap("description")).ColumnWidth = 90
    .Sh.Columns(.ColMap("description")).WrapText = True
    .Sh.Activate
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .ScrollRow = 1
        .ScrollColumn = 1
        .FreezePanes = True
        End With

    End With

For Each DateChar In DateChars.Keys
    Call logMsg(0, "I-004 Date Character [" & DateChar & "] occurrences: " & DateChars(DateChar), "")
    Next DateChar
    
ErrorRows = shPhotosAcc.RowLast - shPhotosCnv.RowLast - IndexRows - BlankRows
    
MsgBox "Conversion finished:" _
    & vbCrLf & "Rows entered (" & (shPhotosAcc.RowLast - 1) & ")" _
    & vbCrLf & "Rows flagged (" & ErrorRows & ")" _
    & vbCrLf & "Rows converted (" & (shPhotosCnv.RowLast - 1) & ")" _
    & vbCrLf & "Rows skipped (INDEXES) (" & IndexRows & ")" _
    & vbCrLf & "Rows skipped (BLANK) (" & BlankRows & ")"
    
Call logMsg(0, "I-005 Rows entered (" & (shPhotosAcc.RowLast - 1) & ")", "")
Call logMsg(0, "I-006 Rows flagged (" & ErrorRows & ")", "")
Call logMsg(0, "I-007 Rows converted (" & (shPhotosCnv.RowLast - 1) & ")", "")
Call logMsg(0, "I-008 Rows skipped (INDEXES) (" & IndexRows & ")", "")
Call logMsg(0, "I-009 Rows skipped (BLANK) (" & BlankRows & ")", "")
    
End Sub

Function initNewData(AccRow As Long) As typNewData
initNewData.AccessRow = AccRow
End Function

Sub addNewData( _
    ShCtl As typShCtl, _
    newData As typNewData)
''''''''''''''''''''''''''

With ShCtl
    .RowLast = .RowLast + 1
    .Sh.Cells(.RowLast, .ColMap("access")) = newData.AccessRow
    .Sh.Cells(.RowLast, .ColMap("library")) = newData.LibNam
    .Sh.Cells(.RowLast, .ColMap("album")) = newData.LibSeq
    .Sh.Cells(.RowLast, .ColMap("pg")) = newData.Pg
    .Sh.Cells(.RowLast, .ColMap("ph")) = newData.Ph
    .Sh.Cells(.RowLast, .ColMap("roll")) = newData.Roll
    .Sh.Cells(.RowLast, .ColMap("city")) = newData.City
    .Sh.Cells(.RowLast, .ColMap("state")) = newData.State
    .Sh.Cells(.RowLast, .ColMap("description")) = newData.Desc
    .Sh.Cells(.RowLast, .ColMap("notes")) = newData.Notes
    .Sh.Cells(.RowLast, .ColMap("date(start)")) = newData.DateStore(newData.DateStoreBest).Date1
    .Sh.Cells(.RowLast, .ColMap("date(end)")) = newData.DateStore(newData.DateStoreBest).Date2
    .Sh.Cells(.RowLast, .ColMap("dr")) = Array("Y", "M", "D")(newData.DateStore(newData.DateStoreBest).DateRange)
    .Sh.Cells(.RowLast, .ColMap("ds")) = Array("(DATE)", "(YR-MON)", "(DESC)", "(WHERE)")(newData.DateStoreBest)
    End With
End Sub


Function parseCandidate( _
    ShCtl As typShCtl, _
    BlankEntries As Long, _
    IndexEntries As Long) _
    As Boolean
''''''''''''''
Dim LibName As String
'''''''''''''''''''''
With ShCtl
 
    Select Case True
    
        Case UCase(Trim(.Sh.Cells(.RowCurr, .ColMap("library")))) = "INDEXES"
        IndexEntries = IndexEntries + 1
        
        Case True _
            And Trim(.Sh.Cells(.RowCurr, .ColMap("mon"))) = "" _
            And Trim(.Sh.Cells(.RowCurr, .ColMap("yr"))) = "" _
            And Trim(.Sh.Cells(.RowCurr, .ColMap("roll"))) = "" _
            And Trim(.Sh.Cells(.RowCurr, .ColMap("date"))) = "" _
            And Trim(.Sh.Cells(.RowCurr, .ColMap("where"))) = "" _
            And Trim(.Sh.Cells(.RowCurr, .ColMap("desc"))) = "" _
            And Trim(.Sh.Cells(.RowCurr, .ColMap("notes"))) = ""
        BlankEntries = BlankEntries + 1
        
        Case Else
        parseCandidate = True
        End Select
    
    End With
End Function

Function parseLibAlbum( _
    ShCtl As typShCtl, _
    newData As typNewData) _
    As Boolean
''''''''''''''
Dim LibName As String
'''''''''''''''''''''
With ShCtl

    LibName = UCase(.Sh.Cells(.RowCurr, .ColMap("library")))
    
    Select Case LibName
    
        Case "BOXES"
        newData.LibNam = "BOXES"
        parseLibAlbum = parseLibAlbumSuffix(ShCtl, "BOX", newData.LibSeq)
            
        Case "BW-NTBK"
        newData.LibNam = "BW-NTBK"
        parseLibAlbum = parseLibAlbumSuffix(ShCtl, "NOTEBK", newData.LibSeq)
            
        Case "BW-NTBK2"
        newData.LibNam = "BW-NTBK"
        parseLibAlbum = parseLibAlbumSuffix(ShCtl, "NOTEBK", newData.LibSeq)
            
        Case "COLORNEG"
        newData.LibNam = "COLORNEG"
        parseLibAlbum = parseLibAlbumSuffix(ShCtl, "NOTEBK", newData.LibSeq)
        
        Case "COLORSLD"
        
        Select Case True
        
            Case .Sh.Cells(.RowCurr, .ColMap("album")) = "DEMOS"
            newData.LibNam = "COLORSLD-1 (DEMO)"
            newData.LibSeq = 1
            parseLibAlbum = True

            Case Mid(.Sh.Cells(.RowCurr, .ColMap("album")), 1, 4) = "DEMO"
            newData.LibNam = "COLORSLD-1 (DEMO)"
            parseLibAlbum = parseLibAlbumSuffix(ShCtl, "DEMO", newData.LibSeq)
 
            Case Mid(.Sh.Cells(.RowCurr, .ColMap("album")), 1, 2) = "CS"
            newData.LibNam = "COLORSLD-2 (CS)"
            parseLibAlbum = parseLibAlbumSuffix(ShCtl, "CS", newData.LibSeq)

            Case Else
            Call logMsg(.RowCurr, "E-010 Unable to parse Album for COLORSLD", _
            .Sh.Cells(.RowCurr, .ColMap("album")))
            End Select
          
        
            
        Case "PLACE"
        newData.LibNam = "PLACE"
        newData.LibSeq = 1
        parseLibAlbum = True
        
        Case "PORTRAIT"
        newData.LibNam = "PORTRAIT"
        parseLibAlbum = parseLibAlbumSuffix(ShCtl, "PORTRT", newData.LibSeq)
            
            
        Case Else
        Call logMsg(.RowCurr, "E-011 Unexpected library name", LibName)
        End Select
    
    End With
End Function

Function parseLibAlbumSuffix( _
    ShCtl As typShCtl, _
    ExpectPfx As String, _
    LibSeq As Long) _
    As Boolean
''''''''''''''
Dim AlbumNam As String
Dim AlbumPfx As String

With ShCtl
AlbumNam = .Sh.Cells(.RowCurr, .ColMap("album"))
AlbumPfx = Mid(AlbumNam, 1, Len(ExpectPfx))

Select Case True

    Case AlbumPfx = ExpectPfx
    
    Select Case True
    
        Case IsNumeric(Mid(AlbumNam, Len(AlbumPfx) + 1))
        LibSeq = Mid(AlbumNam, Len(AlbumPfx) + 1)
        parseLibAlbumSuffix = True
        Exit Function
        
        Case Else
        Call logMsg(.RowCurr, "E-012 Album name suffix not numeric", AlbumNam)
        Exit Function
        End Select
        
    Case Else
    Call logMsg(.RowCurr, "E-013 Album name prefix not [" & ExpectPfx & "]", AlbumNam)
    Exit Function
    End Select
    
End With
End Function

Function parsePgPh( _
    ShCtl As typShCtl, _
    newData As typNewData) _
    As Boolean
''''''''''''''
With ShCtl

    Select Case True
    
        Case Not IsNumeric(.Sh.Cells(.RowCurr, .ColMap("pg")))
        Call logMsg(.RowCurr, "E-014 Non-numeric data in PG column", .Sh.Cells(.RowCurr, .ColMap("pg")))
    
        Case Not IsNumeric(.Sh.Cells(.RowCurr, .ColMap("ph")))
        Call logMsg(.RowCurr, "E-015 Non-numeric data in PH column", .Sh.Cells(.RowCurr, .ColMap("ph")))
        
        Case Else
        newData.Pg = .Sh.Cells(.RowCurr, .ColMap("pg"))
        newData.Ph = .Sh.Cells(.RowCurr, .ColMap("ph"))
        parsePgPh = True
        
        End Select
       
    End With
    
End Function

Function parseRoll( _
    ShCtl As typShCtl, _
    newData As typNewData) _
    As Boolean
''''''''''''''
Dim RollData As String
''''''''''''''''''''''
With ShCtl

    RollData = Trim(.Sh.Cells(.RowCurr, .ColMap("roll")))

    Select Case True
    
        'blank is ok (translates to 0):
        '''''''''''''''''''''''''''''''
        Case RollData = ""
        parseRoll = True
    
        '/S/ is ok (translates to 0):
        '''''''''''''''''''''''''''''
        Case RollData = "S"
        parseRoll = True
        
        Case IsNumeric(RollData)
        newData.Roll = RollData
        parseRoll = True
    
        Case IsNumeric(Mid(RollData, 2))
        newData.Roll = Mid(RollData, 2)
        parseRoll = True
       
        Case Else
        Call logMsg(.RowCurr, "E-016 Non-numeric data in ROLL column", RollData)
        End Select
       
    End With
    
End Function

Function parseWhere( _
    ShCtl As typShCtl, _
    newData As typNewData) _
    As Boolean
''''''''''''''
Dim Where As String
Dim City
Dim State
Dim CityState
'''''''''''''
With ShCtl

    Where = Trim(.Sh.Cells(.RowCurr, .ColMap("where")))
    
    Select Case True
    
        Case Where = ""
        parseWhere = True
        
        Case uTokeniz(Where, CityState, COMMA) <> 1
        Call logMsg(.RowCurr, "E-017 Location is not City-State pair", Where)
        
        Case uTokeniz(CityState(0), City, BLANK) < 0
        Call logMsg(.RowCurr, "E-018 Location is not City-State pair", Where)
        
        Case uTokeniz(CityState(1), State, BLANK) < 0
        Call logMsg(.RowCurr, "E-019 Location is not City-State pair", Where)
        
        Case UBound(State) = 0
        newData.City = uJoin(City, BLANK)
        newData.State = State(0)
        parseWhere = True
        
        Case UBound(State) = 1 And VetDates(.RowCurr, State(1), Null, newData.DateStore(DateStore_WHER))
        newData.City = uJoin(City, BLANK)
        newData.State = State(0)
        parseWhere = True
        
        Case Else
        Call logMsg(.RowCurr, "E-020 Location is not City-State pair", Where)
        End Select
        
    End With
        
End Function

Function parseDescNotes( _
    ShCtl As typShCtl, _
    newData As typNewData) _
    As Boolean
''''''''''''''
Dim Token As Long
Dim LookForDate
''''''''''''''''''''''''''''
With ShCtl
    newData.Desc = Trim(.Sh.Cells(.RowCurr, .ColMap("desc")))
    newData.Notes = Trim(.Sh.Cells(.RowCurr, .ColMap("notes")))
    
    'look for period-delimited date token(s):
    Call uTokeniz(newData.Desc, LookForDate, PERIOD)
    For Token = 0 To UBound(LookForDate)
        If VetDates(.RowCurr, LookForDate(Token), Null, newData.DateStore(DateStore_DESC)) Then
            parseDescNotes = True
            Exit Function
            End If
        Next Token
    
    'look for blank-delimited date token:
    Call uTokeniz(Replace(newData.Desc, PERIOD, BLANK), LookForDate, BLANK)
    For Token = 0 To UBound(LookForDate)
        If VetDates(.RowCurr, LookForDate(Token), Null, newData.DateStore(DateStore_DESC)) Then
            parseDescNotes = True
            Exit Function
            End If
        Next Token
        
    parseDescNotes = True
    End With
    
End Function

Function VetDates(Row As Long, iDate, iDatTokens, oDats As typDateStore) _
    As Boolean
''''''''''''''
'Encode a date string in a DateStore instance
'''''''''''''''''''''''''''''''''''''''''''''
'iDate must be a string that ISDATE will recognize and is the range start date;
'iDatTokens is the date range ENUM (also = # of date components-1) for computing the end date,
'or NULL, in which case iDate is tokenized to arrive at the number dynamically.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'see for more info on ISDATE promiscuity:
'http://spreadsheetpage.com/index.php/tip/understanding_the_isdate_function/
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim DateChar As String * 1
Dim workdate As Date
Dim WorkYear As Long
Dim DatTokens
Dim DatTokenC As Long
Dim I As Long
'''''''''''''

Select Case True

    Case Not IsDate(iDate)
    Exit Function
    
    Case Else
    workdate = CDate(iDate)
    WorkYear = DatePart("yyyy", workdate)
    DatTokenC = IIf( _
            IsNull(iDatTokens), _
            uTokeniz(uReplace(iDate, Array(SLASH, COMMA, HYPHEN), BLANK), DatTokens, BLANK), _
            iDatTokens)
    
    Select Case True
        
        Case WorkYear < 1972 Or WorkYear > currYr
        Call logMsg(Row, "E-021 Bad year in date: " & WorkYear, "")
        Exit Function
        
'        Case Else
'        For I = 1 To Len(iDate)
'            DateChar = Mid(iDate, I, 1)
'            If Not AlphaHash.Exists(DateChar) And Not NumerHash.Exists(DateChar) Then
'                DateChars(DateChar) = DateChars(DateChar) + 1
'                End If
'            Next I
            
        Case DatTokenC < DateRange_Y
        MsgBox "Logic error (1)"
        Exit Function
        
        Case DatTokenC = DateRange_Y
        oDats.Date1 = workdate
        oDats.Date2 = DateAdd("d", -1, DateAdd("yyyy", 1, workdate))
        oDats.DateRange = DateRange_Y
        VetDates = True
            
        Case DatTokenC = DateRange_M
        oDats.Date1 = workdate
        oDats.Date2 = DateAdd("d", -1, DateAdd("m", 1, workdate))
        oDats.DateRange = DateRange_M
        VetDates = True
            
        Case Else
        oDats.Date1 = workdate
        oDats.Date2 = workdate
        oDats.DateRange = DateRange_D
        VetDates = True
        
        End Select
        
    End Select
 
End Function

Function ChooseDate( _
    ShCtl As typShCtl, _
    newData As typNewData) _
    As Boolean
''''''''''''''
Dim I As Long
Dim DateCount As Long
Dim Mon As Long
Dim Yr As Long
Dim DatePartIdx As Long
Dim DatePartStr
Dim DateSourceIdx As Long
Dim DateSourceColNames
Dim MonColData As String
Dim YrColData As String
Dim DateColData As String
Dim DateColMMDDYY As String
'''''''''''''''''''''''''''
DatePartStr = Array("yyyy", "m", "d")

With ShCtl

    DateColData = Trim(.Sh.Cells(.RowCurr, .ColMap("date")))
    MonColData = Trim(.Sh.Cells(.RowCurr, .ColMap("mon")))
    YrColData = Trim(.Sh.Cells(.RowCurr, .ColMap("yr")))
    
    'build date object from DATE column if present:
    '''''''''''''''''''''''''''''''''''''''''''''''
    Select Case Len(DateColData)
    
        Case 0
        'nop
        
        Case 6
        If IsNumeric(DateColData) Then
            DateColMMDDYY = Mid(DateColData, 1, 2) _
            & "-" & Mid(DateColData, 3, 2) _
            & "-" & Mid(DateColData, 5, 2)
            Call VetDates(.RowCurr, DateColMMDDYY, DateRange_D, newData.DateStore(DateStore_DATE))
            End If
            
        Case Else
        Call VetDates(.RowCurr, DateColData, DateRange_D, newData.DateStore(DateStore_DATE))
            
        End Select
        
    'build date object from Yr or Yr/Mon cols:
    ''''''''''''''''''''''''''''''''''''''''''
    Select Case True
            
        Case YrColData <> "" And MonColData <> ""
        Call VetDates(.RowCurr, MonColData & "-" & YrColData, DateRange_M, newData.DateStore(DateStore_YYMM))
        
        Case YrColData <> ""
        Call VetDates(.RowCurr, "1-1-" & YrColData, DateRange_Y, newData.DateStore(DateStore_YYMM))
        
        End Select
        
    End With
    
With newData
        
    'now save the best date so far, checking for inconsistencies:
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For I = LBound(.DateStore) To UBound(.DateStore)
    
        Select Case True
    
            'no date from this source column:
            '''''''''''''''''''''''''''''''''
            Case IsEmpty(.DateStore(I).DateRange)
            
            'first source col with date, save data:
            '''''''''''''''''''''''''''''''''''''''
            Case IsEmpty(.DateStoreBest)
            .DateStoreBest = I
            
            'otherwise check that current date is consistent with previously saved:
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case Else
            'check that date parts in common match:
            '''''''''''''''''''''''''''''''''''''''
            For DatePartIdx = 0 To uMin(.DateStore(.DateStoreBest).DateRange, .DateStore(I).DateRange)
                If DatePart(DatePartStr(DatePartIdx), .DateStore(.DateStoreBest).Date1) _
                <> DatePart(DatePartStr(DatePartIdx), .DateStore(I).Date1) Then
                    Call logMsg(ShCtl.RowCurr, "E-022 Inconsistent dates found", "")
                    Exit Function
                    End If
                Next DatePartIdx
            'dates consistent, save more granular date:
            '''''''''''''''''''''''''''''''''''''''''''
            If .DateStore(.DateStoreBest).DateRange < .DateStore(I).DateRange Then
                .DateStoreBest = I
                End If
            
            End Select
        
        Next I
        
    Select Case True
        
        Case IsEmpty(.DateStoreBest)
        Call logMsg(ShCtl.RowCurr, "E-023 No valid date found", "")
        Exit Function
        
        Case Else
        ChooseDate = True
        
        End Select
        
    End With
    
End Function

Sub logMsg(RowNo As Long, msgTxt As String, ColData)
''''''''''''''''''''''''''''''''''''''''''''''''''''
With shPhotosLog
    .RowLast = .RowLast + 1
    .Sh.Cells(.RowLast, .ColMap("Time")) = Format(Now, "yyyy-mm-dd hhmm")
    .Sh.Cells(.RowLast, .ColMap("Row (Access)")) = RowNo
    .Sh.Cells(.RowLast, .ColMap("Message")) = msgTxt
    .Sh.Cells(.RowLast, .ColMap("Column Data")) = ColData
    End With
End Sub


Function uJoin(iArr, iDlm As String) As String
''''''''''''''''''''''''''''''''''''''''''''''
Dim I As Long
For I = 0 To UBound(iArr)
    If I > 0 Then uJoin = uJoin & iDlm
    uJoin = uJoin & iArr(I)
    Next I
End Function

Function uReplace(iTxt, FindChars, RepChar) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim I As Long
'''''''''''''
uReplace = iTxt
For I = LBound(FindChars) To UBound(FindChars)
    uReplace = Replace(uReplace, FindChars(I), RepChar)
    Next I
End Function

Function uMin(val1, val2)
'''''''''''''''''''''''''
uMin = IIf(val1 < val2, val1, val2)
End Function

Function uMax(val1, val2)
'''''''''''''''''''''''''
uMax = IIf(val1 > val2, val1, val2)
End Function

Function uTokeniz(iTxt As Variant, oArr, iDlm As String) As Long
Dim TmpNext As Long
Dim ResNext As Long
Dim TmpArr
Dim ResArr() As String
''''''''''''''''''''''
TmpArr = Split(iTxt, Mid(iDlm, 1, 1))

For TmpNext = 0 To UBound(TmpArr)
    If TmpArr(TmpNext) <> "" Then
        ReDim Preserve ResArr(0 To ResNext)
        ResArr(ResNext) = TmpArr(TmpNext)
        ResNext = ResNext + 1
        End If
    Next TmpNext
    
Select Case True
    Case ResNext > 0
    oArr = ResArr
    Case Else
    oArr = Array()
    End Select
    
uTokeniz = ResNext - 1
End Function

'REGEX experiments...
'''''''''''''''''''''
Function testdate(datearg As String)
Dim rxResults As Object
Dim Res
Dim Rx As New RegExp
Dim Pfx As String
Pfx = Format(Now, "HH:MM:SS") & " "
Rx.Pattern = "(\d{1,2}\/\d{1,2}\/\d{1,2}\D*)"
Rx.Pattern = "(\d*\/\d*\/\d*)\D*"
Rx.Pattern = "(\d+[\/-]\d+[\/-]\d\d+)"
'Rx.MultiLine = True
Rx.Global = True
Set rxResults = Rx.Execute(datearg)
Debug.Print Pfx & "Count: " & rxResults.Count()


For Each Res In rxResults
    Debug.Print Pfx & "=> " & Res
    Next

End Function

