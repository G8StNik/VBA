Option Explicit
Sub TrimAllCells()

'Trim all cell in a sheet
    Dim cell As Range
    Dim ws As Worksheet
    Dim answer As VbMsgBoxResult
    Set ws = ActiveSheet
    
    answer = MsgBox("Do you want to trim all the cells on this sheet?", vbYesNo + vbQuestion, "Confirm Trim")
    
    If answer = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    
    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) Then
            ' Only trim if the cell contains a string (text)
            If VarType(cell.Value) = vbString Then
                cell.Value = Trim(cell.Value)
            End If
        End If
    Next cell
    
    Application.ScreenUpdating = True
    MsgBox "All cells have been trimmed!"
End Sub
Sub UppercaseAllCells()
    Dim cell As Range
    Dim ws As Worksheet
    Dim answer As VbMsgBoxResult
    Set ws = ActiveSheet

    answer = MsgBox("Do you want to UPPERCASE all letters on this sheet?", vbYesNo + vbQuestion, "Confirm Trim")
    
    If answer = vbNo Then Exit Sub
    Application.ScreenUpdating = False

    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) Then
            If VarType(cell.Value) = vbString Then
                cell.Value = UCase(cell.Value)
            End If
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "All text has been converted to uppercase!"
End Sub
Sub ReplaceBlanksWithNA()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim answer As VbMsgBoxResult
    
    Set ws = ActiveSheet
    Set rng = ws.UsedRange

    answer = MsgBox("Do you want to replace all BLANKS with 'N/A'?", vbYesNo + vbQuestion, "Confirm Trim")
    
    If answer = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    For Each cell In rng
        If IsEmpty(cell.Value) Then
            cell.Value = "N/A"
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "All blank cells in the used range have been replaced with 'N/A'."
End Sub
Sub RemoveDotsFromSelection()
    Dim rng As Range
    
    'Make sure something is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first."
        Exit Sub
    End If
    
    Set rng = Selection
    
    'Replace all "." with nothing in the selection only
    rng.Replace What:=".", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False
                
     
    Application.ScreenUpdating = True
    MsgBox "All periods ('.') have been removed from the selected cells."
End Sub
Sub RemoveCommasFromSelection()
    Dim rng As Range
    
    'Make sure something is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first."
        Exit Sub
    End If
    
    Set rng = Selection
    
    'Replace all "," with nothing in the selection only
    rng.Replace What:=",", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False
                
     
    Application.ScreenUpdating = True
    MsgBox "All commas (',') have been removed from the selected cells."
End Sub
Sub RemoveSemiColonsFromSelection()
    Dim rng As Range
    
    'Make sure something is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first."
        Exit Sub
    End If
    
    Set rng = Selection
    
    'Replace all ";" with nothing in the selection only
    rng.Replace What:=";", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False
                
     
    Application.ScreenUpdating = True
    MsgBox "All semicolons (';') have been removed from the selected cells."
End Sub
Sub RemoveNumberSignsFromSelection()
    Dim rng As Range
    
    'Make sure something is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first."
        Exit Sub
    End If
    
    Set rng = Selection
    
    'Replace all "#" with nothing in the selection only
    rng.Replace What:="#", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False
                
     
    Application.ScreenUpdating = True
    MsgBox "All number signs ('#') have been removed from the selected cells."
End Sub
Sub RemoveDashesFromSelection()
    Dim rng As Range
    
    'Make sure something is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first."
        Exit Sub
    End If
    
    Set rng = Selection
    
    'Replace all "-" with nothing in the selection only
    rng.Replace What:="-", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False
                
     
    Application.ScreenUpdating = True
    MsgBox "All dashes ('-') have been removed from the selected cells."
End Sub
Sub RemoveAsterisksFromSelection()
    Dim rng As Range
    
    'Make sure something is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first."
        Exit Sub
    End If
    
    Set rng = Selection
    
    'Replace all "*" with nothing in the selection only
    rng.Replace What:="~*", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False
                
     
    Application.ScreenUpdating = True
    MsgBox "All asterisks ('*') have been removed from the selected cells."
End Sub
Sub RemoveColonFromSelection()
    Dim rng As Range
    
    'Make sure something is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first."
        Exit Sub
    End If
    
    Set rng = Selection
    
    'Replace all ":" with nothing in the selection only
    rng.Replace What:=":", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False
                
     
    Application.ScreenUpdating = True
    MsgBox "All colons (':') have been removed from the selected cells."
End Sub
Sub RemoveExclamationMarksFromSelection()
    Dim rng As Range
    
    'Make sure something is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first."
        Exit Sub
    End If
    
    Set rng = Selection
    
    'Replace all "!" with nothing in the selection only
    rng.Replace What:="!", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False
                
     
    Application.ScreenUpdating = True
    MsgBox "All exclamation marks ('!') have been removed from the selected cells."
End Sub
Sub RemoveApostropheFromSelection()
    Dim rng As Range
    
    'Make sure something is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first."
        Exit Sub
    End If
    
    Set rng = Selection
    
    'Replace all "'" with nothing in the selection only
    rng.Replace What:="'", Replacement:="", LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False
                
     
    Application.ScreenUpdating = True
    MsgBox "All apostrophes(''') have been removed from the selected cells."
End Sub
Sub NormalizeSpacesAroundDash()
    Dim rng As Range
    Dim cell As Range
    Dim txt As String
    
    ' Make sure something is selected and is a Range
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first."
        Exit Sub
    End If
    
    Set rng = Selection
    
    For Each cell In rng
        If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            txt = cell.Value
            
            ' Replace multiple variants with the standard space-dash-space
            txt = Replace(txt, "  - ", " - ")      ' space space dash space -> space dash space
            txt = Replace(txt, "  -  ", " - ")     ' space space dash space space -> space dash space
            txt = Replace(txt, " -  ", " - ")      ' space dash space space -> space dash space
            
            cell.Value = txt
        End If
    Next cell
    
    MsgBox "Spaces around dashes normalized.", vbInformation
End Sub
Sub ReplaceSingleLetterAmpersand_AllCases()
    Dim rng As Range
    Dim cell As Range
    Dim txt As String
    Dim i As Long
    Dim words() As String
    Dim newWords() As String

    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first."
        Exit Sub
    End If

    Set rng = Selection

    For Each cell In rng
        If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            txt = cell.Value
            txt = Replace(txt, Chr(160), " ") ' Normalize non-breaking spaces
            txt = Replace(txt, "&", " & ")
            words = Split(txt, " ")
            ReDim newWords(0 To UBound(words))
            Dim newIndex As Integer: newIndex = 0
            i = 0
            Do While i <= UBound(words)
                If i <= UBound(words) - 2 Then
                    If Len(words(i)) = 1 And words(i + 1) = "&" And Len(words(i + 2)) = 1 Then
                        newWords(newIndex) = words(i) & "&" & words(i + 2)
                        newIndex = newIndex + 1
                        i = i + 3
                        GoTo NextIteration ' Skip to next loop iteration
                    End If
                End If
                If words(i) <> "" Then
                    newWords(newIndex) = words(i)
                    newIndex = newIndex + 1
                End If
                i = i + 1
NextIteration:
            Loop
            ReDim Preserve newWords(0 To newIndex - 1)
            txt = Join(newWords, " ")
            cell.Value = txt
        End If
    Next cell

    MsgBox "Normalized '&' completed."
End Sub
Option Explicit

Sub AbbreviateWordsUsingDictionary()
    Dim dict As Object
    Dim rng As Range, cell As Range
    Dim word As Variant
    Dim arr As Variant, i As Integer

    ' Create the abbreviation dictionary
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare   ' Case-insensitive lookup

    ' Add abbreviation pairs here
    dict.Add "APARTMENT", "APT"
    dict.Add "AVENUE", "AVE"
    dict.Add "BOULEVARD", "BLVD"
    dict.Add "CIRCLE", "CIR"
    dict.Add "COUNTY", "CO"
    dict.Add "COURT", "CT"
    dict.Add "COVE", "CV"
    dict.Add "DRIVE", "DR"
    dict.Add "EAST", "E"
    dict.Add "HEIGHTS", "HTS"
    dict.Add "HIGHWAY", "HWY"
    dict.Add "LANE", "LN"
    dict.Add "NORTH", "N"
    dict.Add "PARK", "PK"
    dict.Add "PARKWAY", "PKWY"
    dict.Add "PLACE", "PL"
    dict.Add "PLAZA", "PLZ"
    dict.Add "ROAD", "RD"
    dict.Add "SOUTH", "S"
    dict.Add "STREET", "ST"
    dict.Add "SUITE", "STE"
    dict.Add "TRAILER", "TRLR"
    dict.Add "WEST", "W"
    dict.Add "ACCOUNT", "ACCT"
    dict.Add "ACHIEVEMENT", "ACHIEV"
    dict.Add "ADMINISTRATION", "ADMIN"
    dict.Add "ADVANCED", "ADV"
    dict.Add "AIRCONDITIONING", "A/C"
    dict.Add "AMERICA", "AMER"
    dict.Add "ARCHITECTURAL", "ARCH"
    dict.Add "ASSESSOR", "ASR"
    dict.Add "ASSOCIATES", "ASSOC"
    dict.Add "ASSOCIATION", "ASSN"
    dict.Add "ATTORNEYS", "ATTYS"
    dict.Add "AUTHORITY", "AUTH"
    dict.Add "AUTOMOTIVE", "AUTO"
    dict.Add "BOARD", "BRD"
    dict.Add "BROADWAY", "BDWY"
    dict.Add "BROKERAGE", "BKGE"
    dict.Add "BROTHERS", "BROS"
    dict.Add "BUILDER", "BLDR"
    dict.Add "BUILDING", "BLDG"
    dict.Add "BUSINESS", "BUS"
    dict.Add "CENTER", "CTR"
    dict.Add "CLINIC", "CL"
    dict.Add "COLLECTION", "COLLEC"
    dict.Add "COLLECTOR", "COLL"
    dict.Add "COMMERCIAL", "COMMER"
    dict.Add "COMMISSION", "COMM"
    dict.Add "COMMUNITIES", "COMMTS"
    dict.Add "COMMUNITY", "COMM"
    dict.Add "COMPANY", "CO"
    dict.Add "CONCRETE", "CONCR"
    dict.Add "CONSTRUCTION", "CONSTR"
    dict.Add "CONSULTING", "CONS"
    dict.Add "CONTRACTOR", "CONTR"
    dict.Add "CORPORATION", "CORP"
    dict.Add "COVERING", "COV"
    dict.Add "CUSTOM", "CSTM"
    dict.Add "DECORATIVE", "DÉCOR"
    dict.Add "DEPARTMENT", "DEPT"
    dict.Add "DESIGN", "DSGN"
    dict.Add "DEVELOPMENT", "DEV"
    dict.Add "DISBURSEMENT", "DISBURSE"
    dict.Add "DISTRIBUTION", "DISTRIB"
    dict.Add "DISTRICT", "DIST"
    dict.Add "DIVISION", "DIV"
    dict.Add "DOWNTOWN", "DT"
    dict.Add "ECOLOGICAL", "ECOL"
    dict.Add "ELECTRICAL", "ELEC"
    dict.Add "EMBROIDERY", "EMBROID"
    dict.Add "ENFORCEMENT", "ENF"
    dict.Add "ENGINEER", "ENGR"
    dict.Add "ENTERPRISES", "ENT"
    dict.Add "ENVIRONMENTAL", "ENV"
    dict.Add "EQUIPMENT", "EQUIP"
    dict.Add "ESTATE", "EST"
    dict.Add "EXCELLENCE", "EXC"
    dict.Add "EXPORT", "EXP"
    dict.Add "EXPRESS", "EXPR"
    dict.Add "FAMILIES", "FAMS"
    dict.Add "FINANCE", "FIN"
    dict.Add "FINANCIAL", "FIN"
    dict.Add "FLOORING", "FLG"
    dict.Add "FOUNDATION", "FDN"
    dict.Add "GENERAL", "GENL"
    dict.Add "GRAPHICS", "GRAPH"
    dict.Add "GROUP", "GRP"
    dict.Add "HARDWOOD", "HDWD"
    dict.Add "HEATING", "HTG"
    dict.Add "IMPORT", "IMP"
    dict.Add "IMPRESSIONS", "IMPRESS"
    dict.Add "IMPROVEMENT", "IMPROV"
    dict.Add "INDUSTRIAL", "IND"
    dict.Add "INFORMATION", "INFO"
    dict.Add "INSPECTION", "INSP"
    dict.Add "INSTALLATION", "INST"
    dict.Add "INSTITUTIONAL", "INSTL"
    dict.Add "INSURANCE", "INS"
    dict.Add "INTELLIGENCE", "INTEL"
    dict.Add "INTERFACE", "INTF"
    dict.Add "INTERNATIONAL", "INTL"
    dict.Add "INVESTMENTS", "INVEST"
    dict.Add "JANITORIAL", "JAN"
    dict.Add "LICENSE", "LIC"
    dict.Add "LICENSING", "LIC"
    dict.Add "LIGHTING", "LIGHT"
    dict.Add "LIMITED", "LTD"
    dict.Add "LIVING", "LIV"
    dict.Add "MAINTENANCE", "MAINT"
    dict.Add "MANAGEMENT", "MGMT"
    dict.Add "MANAGER", "MGR"
    dict.Add "MATERIALS", "MATS"
    dict.Add "MECHANICAL", "MECH"
    dict.Add "MEDICAL", "MED"
    dict.Add "MILLWORK", "MWK"
    dict.Add "MOBILE", "MOB"
    dict.Add "MOUNTAIN", "MTN"
    dict.Add "NATIONAL", "NATL"
    dict.Add "OCCUPATIONAL", "OCCUP"
    dict.Add "OFFICE", "OFC"
    dict.Add "PAINTING", "PAINT"
    dict.Add "PARKING", "PRKG"
    dict.Add "PLUMBING", "PLUMB"
    dict.Add "POWER", "PWR"
    dict.Add "PRINTING", "PRINT"
    dict.Add "PROFESSIONAL", "PROF"
    dict.Add "PROGRAM", "PROG"
    dict.Add "PROPERTY", "PROP"
    dict.Add "PROTECTION", "PROTECT"
    dict.Add "PUBLIC", "PUB"
    dict.Add "QUARTZ", "QTZ"
    dict.Add "REALTY", "RLTY"
    dict.Add "REDUCTION", "REDUC"
    dict.Add "REGISTER", "REG"
    dict.Add "REGULATION", "REG"
    dict.Add "RELOCATION", "RELO"
    dict.Add "RENTAL", "RNTL"
    dict.Add "RESIDENTIAL", "RES"
    dict.Add "RESTAURANT", "REST"
    dict.Add "RESTORATION", "RESTOR"
    dict.Add "REVENUE", "REV"
    dict.Add "SANITATION", "SANIT"
    dict.Add "SECRETARY", "SEC"
    dict.Add "SECURE", "SEC"
    dict.Add "SECURITY", "SEC"
    dict.Add "SERVICE", "SVC"
    dict.Add "SERVICES", "SVCS"
    dict.Add "SOLUTIONS", "SOLUT"
    dict.Add "SOURCE", "SRC"
    dict.Add "SPECIALTY", "SP"
    dict.Add "STATE", "ST"
    dict.Add "STORAGE", "STOR"
    dict.Add "SUITES", "STES"
    dict.Add "SUPPLY", "SUP"
    dict.Add "SUPPORT", "SUPP"
    dict.Add "SYSTEM", "SYS"
    dict.Add "TECHNOLOGY", "TECH"
    dict.Add "TOWER", "TWR"
    dict.Add "TREASURER", "TREAS"
    dict.Add "UNIVERSAL", "UNIV"
    dict.Add "VEHICLE", "VEH"
    dict.Add "VILLAGE", "VIL"
    dict.Add "VINEYARD", "VNYD"
    dict.Add "VOLUNTARY", "VOL"
    dict.Add "WAREHOUSE", "WHSE"
    dict.Add "WHOLESALE", "WHOL"
    ' Add more as the get approved
    
    ' Ensure range is selected
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range first."
        Exit Sub
    End If
    Set rng = Selection

    replaceCount = 0

    For Each cell In rng
        If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            arr = Split(cell.Value, " ")
            For i = LBound(arr) To UBound(arr)
                word = UCase(arr(i))
                If dict.Exists(word) Then
                    arr(i) = dict(word)
                    replaceCount = replaceCount + 1  ' Increment each time replacement occurs
                End If
            Next i
            cell.Value = Join(arr, " ")
        End If
    Next cell

    MsgBox "Abbreviation Dictionary complete." & vbCrLf & _
           "Words replaced: " & replaceCount, vbInformation
End Sub

Sub ColumnOrder()

    Dim srcWB As Workbook
    Dim srcWS As Worksheet
    Dim tgtWS As Worksheet
    Dim srcPath As String

    Dim srcLastRow As Long
    Dim tgtLastRow As Long
    Dim srcLastCol As Long
    Dim tgtLastCol As Long

    Dim srcHeaders As Object
    Dim i As Long, j As Long
    Dim hdr As String

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '--- update path if needed
    srcPath = "C:\Users\nick.santonato\OneDrive - Artisan Design Group\02_Projects\01_ERP\11_Wave2\KER_RFMS_Customer_part_00000_updated.csv"

    Set tgtWS = ThisWorkbook.Sheets(1) ' WCC_Customer_Remix sheet
    Set srcWB = Workbooks.Open(srcPath)
    Set srcWS = srcWB.Sheets(1)

    Set srcHeaders = CreateObject("Scripting.Dictionary")

    srcLastCol = srcWS.Cells(1, srcWS.Columns.Count).End(xlToLeft).Column
    tgtLastCol = tgtWS.Cells(1, tgtWS.Columns.Count).End(xlToLeft).Column

    '--- build header map from source
    For i = 1 To srcLastCol
        hdr = Trim(srcWS.Cells(1, i).Value)
        If hdr <> "" Then srcHeaders(hdr) = i
    Next i

    srcLastRow = srcWS.Cells(srcWS.Rows.Count, 1).End(xlUp).Row
    tgtLastRow = tgtWS.Cells(tgtWS.Rows.Count, 1).End(xlUp).Row + 1

    '--- copy data by matching headers
    For j = 1 To tgtLastCol
        hdr = Trim(tgtWS.Cells(1, j).Value)

        If srcHeaders.Exists(hdr) Then
            tgtWS.Cells(tgtLastRow, j).Resize(srcLastRow - 1).Value = _
                srcWS.Cells(2, srcHeaders(hdr)).Resize(srcLastRow - 1).Value
        End If
    Next j

    srcWB.Close False

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    MsgBox "RFMS Customer data successfully imported.", vbInformation

End Sub
