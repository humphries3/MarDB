Attribute VB_Name = "mod_app"
Option Explicit
'''''''''''''''

Public Const MOD_BRANCH = "2020-02-23"
Public Const Dbg = True
''''''''''''''''''''''''
'V3 recognize Excel date in DATE COL
''''''''''''''''''''''''''''''''''''

'Noteworthy links:
''''''''''''''''''
'https://docs.microsoft.com/en-us/previous-versions//yab2dx62(v=vs.85)?redirectedfrom=MSDN
'https://stackoverflow.com/questions/40182260/extract-text-using-word-vba-regex-then-save-to-a-variable-as-string
'https://stackoverflow.com/questions/47613786/vba-regex-split-with-variable
'https://stackoverflow.com/questions/10903394/how-to-extract-substring-in-parentheses-using-regex-pattern

Public Const Rgx_MM_YY = "(\D)?(\d{1,2})([-\/])(\d{2})(.)?"
'Public Const Rgx_MM_DD_YY = "\b(\d{1,2})([-\/])(\d{1,2})(\2)(\d{2})\b"
Public Const Rgx_MM_DD_YY = "\b(\d{1,2})([-\/])(\d{1,2})(\2)[ ]?(\d{2})\b"
'Public Const Rgx_MM_DD_YY = "(\d{1,2})([-\/])(\d{1,2})(\2)[ ]?(\d{2})"
Public Const Rgx_MM_DD_YYYY = "\b(\d{1,2})([-\/])(\d{1,2})(\2)(\d{4})\b"
Public Const Rgx_MMDDYY = "\b(\d{2})(\d{2})(\d{2})\b"
Public Const Rgx_YY = "\b(\d{2})\b"
Public Const Rgx_YYYY = "\b(\d{4})\b"
Public Const Rgx_Digits = "^\d+$" 'for verifying digits-only string
Public Const Rgx_Tokens = "\s*(\S+)" 'for parsing whitespace delimited tokens
Public Const Rgx_Spell = _
"\b(january|february|march|april|may|june|july|august|september|october|november|december)[ ]+(\d{1,2})[,]?[ ]?(\d{4})\b"

Public Const COMMA = ","

Public Type typShCtl
    Sh As Worksheet
    RowLast As Long
    RowCurr As Long
    ColMap As New scripting.Dictionary
    End Type
    
Public Type typDateStore
    Date1 As Date
    Date2 As Date
    DateRange As Variant
    End Type
    
'DATERANGE must be variant - IsEmpty used on it
'ENUM over date range types
'''''''''''''''''''''''''''
Public Const DateRange_Y = 0
Public Const DateRange_M = 1
Public Const DateRange_D = 2

'Indexes into DATESTORE array, by date source (no particular order):
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    dateStore(0 To 3) As typDateStore
    DateStoreBest As Variant
    errCnt As Long
    End Type
    
Public shPhotosAcc As typShCtl
Public shPhotosLog As typShCtl
Public shPhotosCnv As typShCtl
Public wbUser As Workbook
Public currYr As Long
Public RegexBlanks As New RegExp
Public RegexDates As New RegExp
Public RgxDigits As New RegExp
Public monDict As New scripting.Dictionary
Public dateStoreDesc
Public dateDlmAllow As New scripting.Dictionary



Sub Run()
'''''''''
Dim ColNo As Long
Dim BlankRows As Long
Dim IndexRows As Long
Dim errCnt As Long
Dim NewColNames
Dim LogColNames
Dim rngAccess As Range
Dim Row As Long
Dim Rng As Range
Dim oLibNam As String
Dim oLibSeq As Long
Dim newData As typNewData
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

' main loop: process each column in each row,
' using a SELECT structure to exit current row handling when a CASE routine returns (F).
' Some routines do not flag data and always return (T) so SELECT cases can proceed.

For shPhotosAcc.RowCurr = 2 To shPhotosAcc.RowLast
''''''''''''''''''''''''''''''''''''''''''''''''''
    'instantiate new row data with i/p seq:
    '''''''''''''''''''''''''''''''''''''''
    newData = initNewData(shPhotosAcc.RowCurr)
    
    Select Case True
    
        'flag rows with INDEX in LIBRARY column or rows missing most data:
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case Not parseCandidate(shPhotosAcc, BlankRows, IndexRows)
        
        'normalize library name and isolate album number in new ALBUM col:
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case Not parseLibAlbum(shPhotosAcc, newData, errCnt)
        
        'validate Pg Ph for numerics:
        '''''''''''''''''''''''''''''
        Case Not parsePgPh(shPhotosAcc, newData, errCnt)
        
        'encode Mon Year if present, always (T):
        ''''''''''''''''''''''''''''''''''''''''
        Case Not parseColMonYr(shPhotosAcc, newData)
        
        'validate/encode roll number (blanks and S -> 0):
        '''''''''''''''''''''''''''''''''''''''''''''''''
        Case Not parseRoll(shPhotosAcc, newData, errCnt)
        
        'encode DATE col if present, always (T):
        ''''''''''''''''''''''''''''''''''''''''
        Case Not parseColDate(shPhotosAcc, newData)
        
        'validate/encode city, state and possible date from WHERE col:
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case Not parseWhere(shPhotosAcc, newData, errCnt)
        
        'parse possible date from DESC col, always (T):
        '''''''''''''''''''''''''''''''''''''''''''''''
        Case Not parseDescNotes(shPhotosAcc, newData)
        
        'ensure some date found, choose best when > 1:
        ''''''''''''''''''''''''''''''''''''''''''''''
        Case Not dateChoose(shPhotosAcc, newData, errCnt)
        
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
    
'ErrorRows = shPhotosAcc.RowLast - shPhotosCnv.RowLast - IndexRows - BlankRows
    
MsgBox "Conversion finished:" _
    & vbCrLf & "Rows entered (" & (shPhotosAcc.RowLast - 1) & ")" _
    & vbCrLf & "Rows flagged (" & errCnt & ")" _
    & vbCrLf & "Rows converted (" & (shPhotosCnv.RowLast - 1) & ")" _
    & vbCrLf & "Rows skipped (INDEXES) (" & IndexRows & ")" _
    & vbCrLf & "Rows skipped (BLANK) (" & BlankRows & ")"
    
Call logMsg(0, "I-005 Rows entered (" & (shPhotosAcc.RowLast - 1) & ")", "")
Call logMsg(0, "I-006 Rows flagged (" & errCnt & ")", "")
Call logMsg(0, "I-007 Rows converted (" & (shPhotosCnv.RowLast - 1) & ")", "")
Call logMsg(0, "I-008 Rows skipped (INDEXES) (" & IndexRows & ")", "")
Call logMsg(0, "I-009 Rows skipped (BLANK) (" & BlankRows & ")", "")
    
End Sub

Sub initRgx()
'''''''''''''
RegexBlanks.Pattern = Rgx_Tokens
RegexBlanks.Global = True
RegexBlanks.MultiLine = False

RegexDates.Global = True
RegexDates.IgnoreCase = True
RegexDates.MultiLine = False

RgxDigits.Global = False
RgxDigits.MultiLine = False
RgxDigits.Pattern = Rgx_Digits
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
    .Sh.Cells(.RowLast, .ColMap("date(start)")) = newData.dateStore(newData.DateStoreBest).Date1
    .Sh.Cells(.RowLast, .ColMap("date(end)")) = newData.dateStore(newData.DateStoreBest).Date2
    .Sh.Cells(.RowLast, .ColMap("dr")) = Array("Y", "M", "D")(newData.dateStore(newData.DateStoreBest).DateRange)
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
    newData As typNewData, _
    errCnt As Long) _
    As Boolean
''''''''''''''
Dim LibName As String
'''''''''''''''''''''
With ShCtl

    LibName = UCase(.Sh.Cells(.RowCurr, .ColMap("library")))
    
    Select Case LibName
    
        Case "BOXES"
        newData.LibNam = "BOXES"
        parseLibAlbum = parseLibAlbumSuffix(ShCtl, "BOX", newData.LibSeq, errCnt)
            
        Case "BW-NTBK"
        newData.LibNam = "BW-NTBK"
        parseLibAlbum = parseLibAlbumSuffix(ShCtl, "NOTEBK", newData.LibSeq, errCnt)
            
        Case "BW-NTBK2"
        newData.LibNam = "BW-NTBK"
        parseLibAlbum = parseLibAlbumSuffix(ShCtl, "NOTEBK", newData.LibSeq, errCnt)
            
        Case "COLORNEG"
        newData.LibNam = "COLORNEG"
        parseLibAlbum = parseLibAlbumSuffix(ShCtl, "NOTEBK", newData.LibSeq, errCnt)
        
        Case "COLORSLD"
        
        Select Case True
        
            Case .Sh.Cells(.RowCurr, .ColMap("album")) = "DEMOS"
            newData.LibNam = "COLORSLD-1 (DEMO)"
            newData.LibSeq = 1
            parseLibAlbum = True

            Case Mid(.Sh.Cells(.RowCurr, .ColMap("album")), 1, 4) = "DEMO"
            newData.LibNam = "COLORSLD-1 (DEMO)"
            parseLibAlbum = parseLibAlbumSuffix(ShCtl, "DEMO", newData.LibSeq, errCnt)
 
            Case Mid(.Sh.Cells(.RowCurr, .ColMap("album")), 1, 2) = "CS"
            newData.LibNam = "COLORSLD-2 (CS)"
            parseLibAlbum = parseLibAlbumSuffix(ShCtl, "CS", newData.LibSeq, errCnt)

            Case Else
            Call logMsg(.RowCurr, "E-010 Unable to parse Album for library COLORSLD", _
            .Sh.Cells(.RowCurr, .ColMap("album")))
            errCnt = errCnt + 1
            End Select
          
        
            
        Case "PLACE"
        newData.LibNam = "PLACE"
        newData.LibSeq = 1
        parseLibAlbum = True
        
        Case "PORTRAIT"
        newData.LibNam = "PORTRAIT"
        parseLibAlbum = parseLibAlbumSuffix(ShCtl, "PORTRT", newData.LibSeq, errCnt)
            
            
        Case Else
        Call logMsg(.RowCurr, "E-011 Unexpected library name", LibName)
        errCnt = errCnt + 1
        End Select
    
    End With
End Function

Function parseLibAlbumSuffix( _
    ShCtl As typShCtl, _
    ExpectPfx As String, _
    LibSeq As Long, _
    errCnt As Long) _
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
    
        Case uDigits(Mid(AlbumNam, Len(AlbumPfx) + 1))
        LibSeq = Mid(AlbumNam, Len(AlbumPfx) + 1)
        parseLibAlbumSuffix = True
        Exit Function
        
        Case Else
        Call logMsg(.RowCurr, "E-012 Album name suffix not numeric", AlbumNam)
        errCnt = errCnt + 1
        Exit Function
        End Select
        
    Case Else
    Call logMsg(.RowCurr, "E-013 Album name prefix not [" & ExpectPfx & "]", AlbumNam)
    errCnt = errCnt + 1
    Exit Function
    End Select
    
End With
End Function

Function parsePgPh( _
    ShCtl As typShCtl, _
    newData As typNewData, _
    errCnt As Long) _
    As Boolean
''''''''''''''
With ShCtl

    Select Case True
    
        Case Not uDigits(.Sh.Cells(.RowCurr, .ColMap("pg")))
        Call logMsg(.RowCurr, "E-014 Non-numeric data in PG column", .Sh.Cells(.RowCurr, .ColMap("pg")))
        errCnt = errCnt + 1
    
        Case Not uDigits(.Sh.Cells(.RowCurr, .ColMap("ph")))
        Call logMsg(.RowCurr, "E-015 Non-numeric data in PH column", .Sh.Cells(.RowCurr, .ColMap("ph")))
        errCnt = errCnt + 1
        
        Case Else
        newData.Pg = .Sh.Cells(.RowCurr, .ColMap("pg"))
        newData.Ph = .Sh.Cells(.RowCurr, .ColMap("ph"))
        parsePgPh = True
        
        End Select
       
    End With
    
End Function

Function parseColMonYr( _
    ShCtl As typShCtl, _
    newData As typNewData) _
    As Boolean
''''''''''''''

With ShCtl

    Select Case True
        
        Case dateFind( _
        .RowCurr, _
        .Sh.Cells(.RowCurr, .ColMap("mon")) & "-" & .Sh.Cells(.RowCurr, .ColMap("yr")), _
        Array("mm-yy"), _
        newData.dateStore, DateStore_YYMM)
        
        Case dateFind( _
        .RowCurr, _
        .Sh.Cells(.RowCurr, .ColMap("yr")), _
        Array("yy"), _
        newData.dateStore, DateStore_YYMM)
         
        End Select
    
    End With
    
'don't reject row even if no date:
parseColMonYr = True
End Function

Function parseRoll( _
    ShCtl As typShCtl, _
    newData As typNewData, _
    errCnt As Long) _
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
        
        Case uDigits(RollData)
        newData.Roll = RollData
        parseRoll = True
    
        Case uDigits(Mid(RollData, 2))
        newData.Roll = Mid(RollData, 2)
        parseRoll = True
       
        Case Else
        Call logMsg(.RowCurr, "E-016 Non-numeric data in ROLL column", RollData)
        errCnt = errCnt + 1
        End Select
       
    End With
    
End Function

Function parseColDate( _
    ShCtl As typShCtl, _
    newData As typNewData) _
    As Boolean
''''''''''''''

With ShCtl

    Select Case True
        
        Case dateFind( _
        .RowCurr, _
        .Sh.Cells(.RowCurr, .ColMap("date")), _
        Array("mmddyy", "mm-dd-yyyy"), _
        newData.dateStore, DateStore_DATE)
        
        Case Else
        'nop
        End Select
    
    End With
    
'don't reject row even if no date:
parseColDate = True
End Function

Function parseWhere( _
    ShCtl As typShCtl, _
    newData As typNewData, _
    errCnt As Long) _
    As Boolean
''''''''''''''
Dim Where As String
Dim City() As String
Dim State() As String
Dim CityState
'''''''''''''

With ShCtl

    Where = .Sh.Cells(.RowCurr, .ColMap("where"))
    
    Select Case True
    
        'if WHERE blank, ok and we're done:
        '''''''''''''''''''''''''''''''''''
        Case Where = ""
        parseWhere = True
        
        'check that we have tokens on either side of COMMA:
        '''''''''''''''''''''''''''''''''''''''''''''''''''
        Case uSplit(Where, CityState, COMMA) <> 1
        Call logMsg(.RowCurr, "E-017 Location is not City-State pair", Where)
        errCnt = errCnt + 1
        
        Case uToken(CityState(0), City) < 1
        Call logMsg(.RowCurr, "E-018 Location is not City-State pair", Where)
        errCnt = errCnt + 1
        
        Case uToken(CityState(1), State) < 1
        Call logMsg(.RowCurr, "E-019 Location is not City-State pair", Where)
        errCnt = errCnt + 1
    
        'if one STATE token, complete cols and finish:
        ''''''''''''''''''''''''''''''''''''''''''''''
        Case UBound(State) = 0
        newData.City = Join(City)
        newData.State = Join(State)
        parseWhere = True
        
        'if date following STATE, encode and strip from STATE col:
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case dateFind(.RowCurr, State(UBound(State)), Array("mm-dd-yyyy", "mmddyy"), newData.dateStore, DateStore_WHER)
        newData.City = Join(City)
        ReDim Preserve State(UBound(State) - 1)
        newData.State = Join(State)
        parseWhere = True
        
        Case Else
        'multiple STATE tokens but no date, use all tokens for STATE:
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        newData.City = Join(City)
        newData.State = Join(State)
        parseWhere = True
        End Select
        
    End With
        
End Function

Function parseDescNotes( _
    ShCtl As typShCtl, _
    newData As typNewData) _
    As Boolean
''''''''''''''
With ShCtl
    
    'CASE structure here is a formality - in case other paths added in future
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case True
    
        Case dateFind(.RowCurr, _
            .Sh.Cells(.RowCurr, .ColMap("desc")), _
            Array("mm-dd-yy", "spell", "mm-yy"), _
            newData.dateStore, DateStore_DESC)
            
        Case Else
        'nop
            
        End Select
    
    End With
    
parseDescNotes = True
End Function

Function dateFind( _
    Row As Long, _
    iStr As String, _
    iFormats, _
    oDats() As typDateStore, _
    dateStorIdx) _
    As Boolean
''''''''''''''
'FIND A DATE in a string using regex patterns from caller.
'ISTR defined as STRING to cast date datatype from spreadsheet cell that may be DATE not character
'IFORMATS is an array a regex patterns.
'ODATS is a DATASTORE entry for encoding date.
'Logic: search string for a valid date using regex patterns in order supplied.
'First valid date ends process.
'Caller responsible for coding patterns in order of decreasing granularity so most specific date chosen.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim workYear As Long
Dim DatTokens
Dim DatTokenC As Long
Dim I As Long
Dim iFormatC As Integer
Dim errDlmPos As Integer
Dim iFormatW As String
Dim matchList As MatchCollection
Dim MatchC As Integer
'''''''''''''''''''''
If Dbg Then Call logMsg( _
    Row, _
    "D-01 DF (" & dateStoreDesc(dateStorIdx) & ")", _
    "[" _
    & iStr _
    & "]")
    
For iFormatC = 0 To UBound(iFormats)
    iFormatW = iFormats(iFormatC)
    
    Select Case iFormatW
    
        Case "spell"
        RegexDates.Pattern = Rgx_Spell
        Set matchList = RegexDates.Execute(iStr)
        For MatchC = 1 To matchList.Count
            If Dbg Then Call DfLogMsg(Row, matchList(MatchC - 1), iFormatW)
        
            Select Case True
            
                Case Not monDict.Exists(matchList(MatchC - 1).SubMatches(0))
                
                Case dateValid( _
                Row, _
                monDict(matchList(MatchC - 1).SubMatches(0)), _
                matchList(MatchC - 1).SubMatches(1), _
                matchList(MatchC - 1).SubMatches(2), _
                oDats(dateStorIdx).Date1)
                oDats(dateStorIdx).DateRange = DateRange_D
                oDats(dateStorIdx).Date2 = oDats(dateStorIdx).Date1
                dateFind = True
                Exit Function
                
                End Select
                
             Next MatchC
    
        Case "mm-dd-yy"
        RegexDates.Pattern = Rgx_MM_DD_YY
        Set matchList = RegexDates.Execute(iStr)
        For MatchC = 1 To matchList.Count
            If Dbg Then Call DfLogMsg(Row, matchList(MatchC - 1), iFormatW)
            If dateValid( _
            Row, _
            matchList(MatchC - 1).SubMatches(0), _
            matchList(MatchC - 1).SubMatches(2), _
            matchList(MatchC - 1).SubMatches(4), _
            oDats(dateStorIdx).Date1) Then
                oDats(dateStorIdx).DateRange = DateRange_D
                oDats(dateStorIdx).Date2 = oDats(dateStorIdx).Date1
                dateFind = True
                Exit Function
                End If
            Next MatchC
            
        Case "mm-dd-yyyy"
        RegexDates.Pattern = Rgx_MM_DD_YYYY
        Set matchList = RegexDates.Execute(iStr)
        For MatchC = 1 To matchList.Count
            If Dbg Then Call DfLogMsg(Row, matchList(MatchC - 1), iFormatW)
            If dateValid( _
            Row, _
            matchList(MatchC - 1).SubMatches(0), _
            matchList(MatchC - 1).SubMatches(2), _
            matchList(MatchC - 1).SubMatches(4), _
            oDats(dateStorIdx).Date1) Then
                oDats(dateStorIdx).DateRange = DateRange_D
                oDats(dateStorIdx).Date2 = oDats(dateStorIdx).Date1
                dateFind = True
                Exit Function
                End If
            Next MatchC
            
        Case "mm-yy"
        RegexDates.Pattern = Rgx_MM_YY
        Set matchList = RegexDates.Execute(iStr)
        errDlmPos = -1
        

        For MatchC = 1 To matchList.Count
            If Dbg Then Call DfLogMsg(Row, matchList(MatchC - 1), iFormatW)
            
            Select Case True
                 
                'If we picked up an improperly isolated mm-yy with our regex, skip:
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Case errDlmPos = matchList(MatchC - 1).FirstIndex
                 
                'If we picked up an improperly isolated mm-yy with our regex, skip:
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                Case Not dateDlmAllow.Exists(matchList(MatchC - 1).SubMatches(0))
               
                Case Not dateDlmAllow.Exists(matchList(MatchC - 1).SubMatches(4))
                errDlmPos = matchList(MatchC - 1).FirstIndex + matchList(MatchC - 1).Length
                  
                'otherwise attempt validation:
                ''''''''''''''''''''''''''''''
                Case dateValid( _
                Row, _
                matchList(MatchC - 1).SubMatches(1), _
                1, _
                matchList(MatchC - 1).SubMatches(3), _
                oDats(dateStorIdx).Date1)
                oDats(dateStorIdx).DateRange = DateRange_M
                oDats(dateStorIdx).Date2 = DateAdd("d", -1, DateAdd("m", 1, oDats(dateStorIdx).Date1))
                dateFind = True
                Exit Function
                End Select
            
            Next MatchC
   
        Case "mmddyy"
        RegexDates.Pattern = Rgx_MMDDYY
        Set matchList = RegexDates.Execute(iStr)
        For MatchC = 1 To matchList.Count
            If Dbg Then Call DfLogMsg(Row, matchList(MatchC - 1), iFormatW)
            If dateValid( _
            Row, _
            matchList(MatchC - 1).SubMatches(0), _
            matchList(MatchC - 1).SubMatches(1), _
            matchList(MatchC - 1).SubMatches(2), _
            oDats(dateStorIdx).Date1) Then
                oDats(dateStorIdx).DateRange = DateRange_D
                oDats(dateStorIdx).Date2 = oDats(dateStorIdx).Date1
                dateFind = True
                Exit Function
                End If
            Next MatchC
   
        Case "yyyy"
        RegexDates.Pattern = Rgx_YYYY
        Set matchList = RegexDates.Execute(iStr)
        For MatchC = 1 To matchList.Count
            If Dbg Then Call DfLogMsg(Row, matchList(MatchC - 1), iFormatW)
            If dateValid( _
            Row, _
            1, _
            1, _
            matchList(MatchC - 1).SubMatches(0), _
            oDats(dateStorIdx).Date1) Then
                oDats(dateStorIdx).Date2 = DateAdd("d", -1, DateAdd("yyyy", 1, oDats(dateStorIdx).Date1))
                oDats(dateStorIdx).DateRange = DateRange_Y
                dateFind = True
                Exit Function
                End If
            Next MatchC
            
        Case "yy"
        RegexDates.Pattern = Rgx_YY
        Set matchList = RegexDates.Execute(iStr)
        For MatchC = 1 To matchList.Count
            If Dbg Then Call DfLogMsg(Row, matchList(MatchC - 1), iFormatW)
            If dateValid( _
            Row, _
            1, _
            1, _
            matchList(MatchC - 1).SubMatches(0), _
            oDats(dateStorIdx).Date1) Then
                oDats(dateStorIdx).Date2 = DateAdd("d", -1, DateAdd("yyyy", 1, oDats(dateStorIdx).Date1))
                oDats(dateStorIdx).DateRange = DateRange_Y
                dateFind = True
                Exit Function
                End If
            Next MatchC
            
        Case Else
        MsgBox "Unknown date format [" & iFormatW & "]", vbCritical
        Error 100
        Exit Function
        
        End Select
        
    Next iFormatC
            
End Function

Function DfLogMsg(Row As Long, argMatch As Match, iFormat As String)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim I As Integer
Dim subListStr As String
''''''''''''''''''''''''
For I = 0 To argMatch.SubMatches.Count - 1
    Select Case True
        Case I = 0
        subListStr = "[" & argMatch.SubMatches(I) & "]"
        Case Else
        subListStr = subListStr & " [" & argMatch.SubMatches(I) & "]"
        End Select
        
    Next I

Call logMsg(Row, "D-02 Rgx [" & iFormat & "]", subListStr)
    
End Function

Function dateValid(Row As Long, MM, DD, YY, oDat As Date) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim workDate As String
Dim workYear As Integer
'''''''''''''''''''''''
workDate = MM & "/" & DD & "/" & Format(YY, "00")

Select Case True

    Case Not IsDate(workDate)
    If Dbg Then Call dateValidLog(Row, MM, DD, YY, "bad date")
    Exit Function
    
    Case Else
    oDat = CDate(workDate)
    workYear = Year(oDat)
    
    Select Case True
    
        Case workYear < 1980 Or workYear > currYr
        'Call logMsg(Row, "E-021 Bad year in date: " & workYear, workDate)
        If Dbg Then Call dateValidLog(Row, MM, DD, YY, "bad year")
        Exit Function
        
        Case Else
        dateValid = True
        If Dbg Then Call dateValidLog(Row, MM, DD, YY, "ok")
        Exit Function
        End Select
    
    End Select
    
End Function

Sub dateValidLog(Row As Long, MM, DD, YY, Status)
Call logMsg( _
    Row, _
    "D-02 DV", _
    "[" & MM & "] [" & DD & "] [" & YY & "] [" & Status & "]")
End Sub

Function dateChoose( _
    ShCtl As typShCtl, _
    newData As typNewData, _
    errCnt As Long) _
    As Boolean
''''''''''''''
Dim I As Long
Dim DatePartIdx As Long
Dim DatePartStr
'''''''''''''''

DatePartStr = Array("yyyy", "m", "d")

With newData
        
    'now save the best date so far, checking for inconsistencies:
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For I = LBound(.dateStore) To UBound(.dateStore)
    
        Select Case True
    
            'no date from this source column:
            '''''''''''''''''''''''''''''''''
            Case IsEmpty(.dateStore(I).DateRange)
            
            'first source col with date, save data:
            '''''''''''''''''''''''''''''''''''''''
            Case IsEmpty(.DateStoreBest)
            .DateStoreBest = I
            
            'otherwise check that current date is consistent with previously saved:
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Case Else
            'check that date parts in common match:
            '''''''''''''''''''''''''''''''''''''''
            For DatePartIdx = 0 To uMin(.dateStore(.DateStoreBest).DateRange, .dateStore(I).DateRange)
                If DatePart(DatePartStr(DatePartIdx), .dateStore(.DateStoreBest).Date1) _
                <> DatePart(DatePartStr(DatePartIdx), .dateStore(I).Date1) Then
                    Call logMsg(ShCtl.RowCurr, "E-022 Inconsistent dates found", "")
                    errCnt = errCnt + 1
                    Exit Function
                    End If
                Next DatePartIdx
                
            'dates consistent, save more granular date:
            '''''''''''''''''''''''''''''''''''''''''''
            If .dateStore(.DateStoreBest).DateRange < .dateStore(I).DateRange Then
                .DateStoreBest = I
                End If
            
            End Select
        
        Next I
        
    Select Case True
        
        Case IsEmpty(.DateStoreBest)
        Call logMsg(ShCtl.RowCurr, "E-023 No valid date found", "")
        errCnt = errCnt + 1
        Exit Function
        
        Case Else
        dateChoose = True
        
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


Function uMin(val1, val2)
'''''''''''''''''''''''''
uMin = IIf(val1 < val2, val1, val2)
End Function

Function uMax(val1, val2)
'''''''''''''''''''''''''
uMax = IIf(val1 > val2, val1, val2)
End Function

Function uDigits(iStr As String) As Boolean
'''''''''''''''''''''''''''''''''''''''''''
uDigits = (RgxDigits.Execute(iStr).Count > 0)
End Function

Function uSplit( _
    iTxt As String, _
    oArr, _
    iDlm As String) _
    As Integer
''''''''''''''
oArr = Split(iTxt, iDlm)
uSplit = UBound(oArr)
End Function
     
Function uToken(iTxt, oArr() As String) As Integer
Dim Matches As MatchCollection
Dim I As Integer
''''''''''''''''
Set Matches = RegexBlanks.Execute(iTxt)

If Matches.Count > 0 Then
    ReDim oArr(Matches.Count - 1)
    For I = 1 To Matches.Count
        oArr(I - 1) = Matches(I - 1).SubMatches(0)
        Next I
    End If
    
uToken = Matches.Count
End Function

Function uTokenMC(iTxt As String, MC As MatchCollection) As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Experimental: avoid array build expense by tokenizing directly into
' caller MC. Doesn't work for us since we would incur same overhead rebuilding strings
' from MC
'''''''''
Dim Regex As New RegExp
'''''''''''''''''''''''
Regex.Pattern = "\s*(\S+)"
Regex.Global = True
Set MC = Regex.Execute(iTxt)
uTokenMC = MC.Count
End Function

Function zzTokenMC(Txt1 As String, Txt2 As String)
Dim I As Integer
Dim Arr1 As MatchCollection
Dim Arr2 As MatchCollection
'''''''''''''''''''''''''''
Select Case True

    Case uTokenMC(Txt1, Arr1) = 0 Or uTokenMC(Txt2, Arr2) = 0
    Debug.Print "missing STRS"
    
    
    Case Else
    
    For I = 0 To Arr1.Count - 1
        Debug.Print "1: " & Arr1(I).SubMatches(0)
        Next I
        
    For I = 0 To Arr2.Count - 1
        Debug.Print "2: " & Arr2(I).SubMatches(0)
        Next I
        
    End Select
        
End Function

'REGEX experiments...
'''''''''''''''''''''
Function zzTestdate(datearg As String)
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

Function zzNumer(iTxt As String)
Dim myRegex As New RegExp
Dim I

myRegex.Pattern = "^\d+$"

Debug.Print "S1: " & myRegex.Execute(iTxt).Count
Exit Function

Dim Matches As MatchCollection
Dim Match As Match

Set Matches = myRegex.Execute(iTxt)
For I = 1 To Matches.Count
    Debug.Print "S2: " & Matches(I - 1).SubMatches(0)
    Next I
    
Debug.Print "S2: " & Matches.Count
Debug.Print "---" & vbCrLf
End Function

Function zzDat(iDat As Date) As String
''''''''''''''''''''''''''''''''''''''
' Experiment: show that date cast to string returns m/d/yyyy
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
zzDat = iDat
End Function

Function zzToks(rxSel As String, Param As String)
'''''''''''''''''''''''''''''''''''''''''''''''''
'Experiment: pass different selector characters in regex pattern
'rxSel is regex non-whitespace selector character, (S|w) probably...
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Matches As MatchCollection
Dim Rx As New RegExp
Dim Pfx As String
Dim I As Integer
Dim Arr() As String

Select Case True

    Case Len(rxSel) <> 1
    zzToks = "Invalid regex selector char"
    Exit Function
    
    Case Else
    Pfx = Format(Now, "HH:MM:SS") & " "
    Rx.Pattern = "\s*(\" & rxSel & "+)"
    Rx.Global = True
    Set Matches = Rx.Execute(Param)
    
    If Matches.Count > 0 Then
        ReDim Arr(Matches.Count - 1)
    
        For I = 1 To Matches.Count
            Debug.Print Pfx & "=> [" & Matches(I - 1).SubMatches(0) & "]"
            Arr(I - 1) = Matches(I - 1).SubMatches(0)
            Next
            
        Debug.Print Pfx & "(J1) [" & Join(Arr) & "]"
        
        If UBound(Arr) > 0 Then
            ReDim Preserve Arr(UBound(Arr) - 1)
            Debug.Print Pfx & "(J2) [" & Join(Arr) & "]"
            End If
            
        End If
        
    zzToks = Pfx & "Count: " & Matches.Count & vbCrLf
    End Select
    
End Function

Function zzFormat(Param As Integer)
'''''''''''''''''''''''''''''''''''
' Experiment: show that '0' mask indicates min leading digits
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Debug.Print Format(Param, "00")
End Function

