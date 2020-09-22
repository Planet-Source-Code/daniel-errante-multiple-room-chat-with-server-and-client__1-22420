Attribute VB_Name = "iniread"
#If Win16 Then


Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal filename As String) As Integer


Declare Function GetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfilestring" (ByVal AppName As String, ByVal KeyName As Any, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal filename As String) As Integer
#Else
    ' NOTE: The lpKeyName argument for GetPr
    '     ofileString, WriteProfileString,
    'GetPrivateProfileString, and WritePriva
    '     teProfileString can be either
    'a string or NULL. This is why the argum
    '     ent is defined as "As Any".
    ' For example, to pass a string specifyB
    '     yVal "wallpaper"
    ' To pass NULL specifyByVal 0&
    'You can also pass NULL for the lpString
    '     argument for WriteProfileString
    'and WritePrivateProfileString
    ' Below it has been changed to a string
    '     due to the ability to use vbNullString


Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If
Public THEFILE As String

Function ReadIni2(strSectionHeader As String, strVariableName As String) As String
    '*** DESCRIPTION:Reads from an *.INI fil
    '     e strFileName (full path & file name)
    '*** RETURNS:The string stored in [strSe
    '     ctionHeader], line beginning strVariableName=
    '*** NOTE: Requires declaration of API c
    '     all GetPrivateProfileString
    'Initialise variable
    Dim strReturn As String
    'Blank the return string
    strReturn = String(255, Chr(0))
    'Get requested information, trimming the
    '     returned string
    ReadIni2 = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), THEFILE))
    Exit Function
End Function


Function ParseString(strIn As String, intOffset As Integer, strDelimiter As String) As String
    '*** DESCRIPTION:Parses the passed strin
    '     g, returning the value indicated
    '***by the offset specified, eg: the str
    '     ing "Hello, World ","
    '***offset 2 = "World".
    '*** RETURNS:See description.
    '*** NOTE: The offset starts at 1 and th
    '     e delimiter is the Character
    '***which separates the elements of the
    '     string.
    'Trap any bad calls


    If Len(strIn) = 0 Or intOffset = 0 Then
        ParseString = ""
        Exit Function
    End If
    'Declare local variables
    Dim intStartPos As Integer
    ReDim intDelimPos(10) As Integer
    Dim intStrLen As Integer
    Dim intNoOfDelims As Integer
    Dim intCount As Integer
    Dim strQuotationMarks As String
    Dim intInsideQuotationMarks As Integer
    strQuotationMarks = Chr(34) & Chr(147) & Chr(148)
    intInsideQuotationMarks = False


    For intCount = 1 To Len(strIn)
        'If character is a double-quote then tog
        '     gle the In Quotation flag


        If InStr(strQuotationMarks, Mid$(strIn, intCount, 1)) <> 0 Then
            intInsideQuotationMarks = (Not intInsideQuotationMarks)
        End If
        If (Not intInsideQuotationMarks) And (Mid$(strIn, intCount, 1) = strDelimiter) Then
        intNoOfDelims = intNoOfDelims + 1
        'If array filled then enlarge it, keepin
        '     g existing contents


        If (intNoOfDelims Mod 10) = 0 Then
            ReDim Preserve intDelimPos(intNoOfDelims + 10)
        End If
        intDelimPos(intNoOfDelims) = intCount
    End If
Next intCount
'Handle request for value not present (o
'     ver-run)


If intOffset > (intNoOfDelims + 1) Then
    ParseString = ""
    Exit Function
End If
'Handle boundaries of string


If intOffset = 1 Then
    intStartPos = 1
End If
'Requesting last value - handle null


If intOffset = (intNoOfDelims + 1) Then


    If Right$(strIn, 1) = strDelimiter Then
        intStartPos = -1
        intStrLen = -1
        ParseString = ""
        Exit Function
    Else
        intStrLen = Len(strIn) - intDelimPos(intOffset - 1)
    End If
End If
'Set start and length variables if not h
'     andled by boundary check above


If intStartPos = 0 Then
    intStartPos = intDelimPos(intOffset - 1) + 1
End If


If intStrLen = 0 Then
    intStrLen = intDelimPos(intOffset) - intStartPos
End If
'Set the return string
ParseString = Mid$(strIn, intStartPos, intStrLen)
End Function


Function WriteIni(strSectionHeader As String, strVariableName As String, strValue As String) As Integer
    
    '*** DESCRIPTION:Writes to an *.INI file
    '     called strFileName (full path & file name)
    '*** RETURNS:Integer indicating failure
    '     (0) or success (other) To write
    '*** NOTE: Requires declaration of API c
    '     all WritePrivateProfileString
    'Call the API
    WriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, THEFILE)
End Function
