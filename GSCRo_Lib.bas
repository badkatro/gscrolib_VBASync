Attribute VB_Name = "GSCRo_Lib"
'Option Base 1
'
Type ExtractedNumbers
    count As Integer
    numbers() As String
End Type
'______________________________________________________________________
'
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszpath As String) As Long
'
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long
'
Public Type BrowseInfo
    hOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lparam As Long
    iImage As Long
End Type

Public Const MProdBaseFolder As String = "M:\PROD\"
'
Public Const GWL_STYLE = (-16)      ' General style bits of windows offset
Public Const GWL_EXSTYLE = (-20)    ' Extended sytle bits of windows offset
'Public Const BIF_USENEWUI = &H40    ' Constant for using new interface for SHBrowseForFolder browse dialog api function
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type foundFilesStruct
    count As Integer
    files() As String
    sizeInKb As Long
End Type

Public Type foundFoldersStruct
    count As Integer
    folders() As String
End Type

'____________________________________________________________________________________________________________________
'
' Destroys the specified window and all its child windows
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
' Finds the first top level window in the window list that meets the specified conditions
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal WindowName As String) As Long
' Obtains the handle of the active window
Public Declare Function GetActiveWindow Lib "user32" () As Long
'
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' Gets Environment Variable, impractical to use because of its parameters and return type
Public Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, _
ByVal lpBuffer As String, ByVal nSize As Long) As Long
' Obtains information from the window structure for the specified window
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
' Determines if given window handle is valid
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
' Change position and size of the specified window. Size parameters may be overridden by min and max settings for top level windows
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal bRepaint As Long) As Long     ' x, y - new coordinates, nWidth, nHeight - new dimensions, bRepaint - nonzero to redraw, 0 otherwise
'
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
' Sets an environment variable, valable for current Windows session
Public Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
' Sets information in the window structure for the specified window
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lparam As Any) As Long
'_____________________________________________________________________________________________________________________
'
Private BaseDocName As String                       ' BackupQ
Private DocType(1, 21) As String                    ' BackupQ
Private DocSuff(1, 5) As String                     ' BackupQ
Private DocYear() As String                         ' BackupQ
Private EUDocLang(1, 24) As String                  ' BackupQ
Public SGCDocOrig As String, SGCDocRo As String     ' BackupQ
Public racounter As Integer                         ' ListAllSubfolders, GetAllSubfolders
Public Const DpCnt As String = "c:\gsc6\Word\Cont\Lro\"
Public Const ProjectsBaseFolder As String = "T:\Prepared originals\"


'*************************************************************************************************************************************************************************************************
'*************************************************************************************************************************************************************************************************
'
'These functions need the Registry key with its complete path, so i_RegKey must always begin with one of the following values:
'•HKCU or HKEY_CURRENT_USER
'•HKLM or HKEY_LOCAL_MACHINE
'•HKCR or HKEY_CLASSES_ROOT
'•HKEY_USERS
'•HKEY_CURRENT_CONFIG
'and end with the name of the key...
'
'RegKeySave also has an input parameter for the type of the Registry key value. Supported are the following types:•REG_SZ - A string. If the type is not specified, this will be used as Default.
'•REG_DWORD - A 32-bit number.
'•REG_EXPAND_SZ - A string that contains unexpanded references to environment variables.
'•REG_BINARY - Binary data in any form. You really shouldn't touch such entries.
'You can find more info about Registry value types in the MSDN.
'
'*************************************************************************************************************************************************************************************************
'
'
'reads the value for the registry key i_RegKey
'if the key cannot be found, the return value is ""
Function RegKeyRead(i_RegKey As String) As String

Dim myWS As Object

    On Error Resume Next
    'access Windows scripting
    Set myWS = CreateObject("WScript.Shell")
    'read key from registry
    RegKeyRead = myWS.RegRead(i_RegKey)


    Set myWS = Nothing
End Function

'returns True if the registry key i_RegKey was found
'and False if not
Function RegKeyExists(i_RegKey As String) As Boolean
Dim myWS As Object

  On Error GoTo ErrorHandler
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'try to read the registry key
  myWS.RegRead i_RegKey
  'key was found
  RegKeyExists = True
  Exit Function
  
ErrorHandler:
  'key was not found
  RegKeyExists = False
End Function

'sets the registry key i_RegKey to the
'value i_Value with type i_Type
'if i_Type is omitted, the value will be saved as string
'if i_RegKey wasn't found, a new registry key will be created
Sub RegKeySave(i_RegKey As String, i_Value As String, Optional i_Type As String = "REG_SZ")
Dim myWS As Object

  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'write registry key
  myWS.RegWrite i_RegKey, i_Value, i_Type

End Sub



'deletes i_RegKey from the registry
'returns True if the deletion was successful,
'and False if not (the key couldn't be found)
Function RegKeyDelete(i_RegKey As String) As Boolean
Dim myWS As Object

  On Error GoTo ErrorHandler
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'delete registry key
  myWS.RegDelete i_RegKey
  'deletion was successful
  RegKeyDelete = True
  Exit Function

ErrorHandler:
  'deletion wasn't successful
  RegKeyDelete = False
End Function
'
'*****************************************************************************************************************************************
'


Function getProjectFolderOf(documentName As String) As String   ' should return folder object reference or string (complete path?)


Dim fso As New Scripting.FileSystemObject
Dim ctDocLanguage As String
Dim baseName As String





If Is_GSC_Doc(documentName) Then
        
        ctDocLanguage = getLanguage_fromDocName(documentName)
        
        If Right$(documentName, 4) = ".doc" Or Right$(documentName, 5) = ".docx" Or _
            Right$(documentName, 4) = ".rtf" Or Right$(documentName, 4) = ".txt" Then
            baseName = getGSCBaseName(documentName)
            
        ElseIf LCase(Right(documentName, 5)) = ".copy" Then
            baseName = fso.GetBaseName(documentName)
            
            If Right$(baseName, 4) = ".doc" Or Right$(baseName, 5) = ".docx" Then
                baseName = fso.GetBaseName(baseName)
            End If
            
        Else
            baseName = documentName
        End If
        
        If ctDocLanguage <> "" Then
            
            Select Case ctDocLanguage
                
                Case "en"
                    
                    If fso.FolderExists(ProjectsBaseFolder & baseName) Then
                        getProjectFolderOf = fso.GetFolder(ProjectsBaseFolder & baseName).path
                        Exit Function
                    Else
                        getProjectFolderOf = ""
                        Exit Function
                    End If
                                                            
                Case "fr"
                    
                    If fso.FolderExists(ProjectsBaseFolder & baseName) Then
                        getProjectFolderOf = fso.GetFolder(ProjectsBaseFolder & baseName).path
                        Exit Function
                    ElseIf fso.FolderExists(ProjectsBaseFolder & Replace(baseName, ctDocLanguage, "en")) Then
                        getProjectFolderOf = fso.GetFolder(ProjectsBaseFolder & Replace(baseName, ctDocLanguage, "en")).path
                        Exit Function
                    Else
                        getProjectFolderOf = ""
                        Exit Function
                    End If
                    
                Case Else
                    
                    If fso.FolderExists(ProjectsBaseFolder & baseName) Then
                        getProjectFolderOf = fso.GetFolder(ProjectsBaseFolder & baseName).path
                        Exit Function
                    ElseIf fso.FolderExists(ProjectsBaseFolder & Replace(baseName, ctDocLanguage, "en")) Then
                        getProjectFolderOf = fso.GetFolder(ProjectsBaseFolder & Replace(baseName, ctDocLanguage, "en")).path
                        Exit Function
                    ElseIf fso.FolderExists(ProjectsBaseFolder & Replace(baseName, ctDocLanguage, "fr")) Then
                        getProjectFolderOf = fso.GetFolder(ProjectsBaseFolder & Replace(baseName, ctDocLanguage, "fr")).path
                        Exit Function
                    Else
                        getProjectFolderOf = ""
                        Exit Function
                    End If
                    
            End Select
            
        Else
            getProjectFolderOf = ""
            Debug.Print "getProjectFolderOf: getLanguage_fromDocName(" & documentName & ") resulted in empty string value!"
            Exit Function
        End If
        
Else
    getProjectFolderOf = ""
    Exit Function
End If

Set fso = Nothing
End Function

Function GetOriginalLanguage_fromMProd(GSCFilename As String) As String
' Works. To complete implementation by reading DW metadata when present instead of this.
' Still useful for when we have no file opened (it accepts a file name as input)

Dim originalLang As String
Dim mprodOriginalPath As String

Dim unitlg As String
unitlg = LCase(Mid$(Environ("computername"), 3, 2))

Dim mprodOriginalFileName As String

If Is_GSC_Doc(GSCFilename) Then
        
    ' Need a Is_DW_Document boolean function which should attempt extraction of DW metadata
    ' Need a Get_DW_OriginalLanguage string function to get the or lang of current RO document from metadata if DW (using the above)
    
    ' If not DW, prob worth trying the old method of circling the opened files to find original language in the (probably) opened original
    ' or even measure the length it takes to bring same using the above or the below method, involving querying network
        
    mprodOriginalPath = Build_MProd_Path(GSCFileName_ToStandardName(GSCFilename))
    mprodOriginalFileName = Build_MProd_Path(GSCFileName_ToStandardName(GSCFilename), "getFileNameOnly")
    
    If mprodOriginalPath <> "" Then
        GetOriginalLanguage_fromMProd = getLanguage_fromDocName(mprodOriginalFileName)
    Else    ' No file was found on M Prod named as sourcedocument
        GetOriginalLanguage_fromMProd = ""      ' admittance of failure (doc was not on MProd, as specified)
    End If
    
    
Else
    StatusBar = "Get Original Language: Current document NOT Council, Exiting..."
    Exit Function
End If

End Function

Function getDocType_fromDocName(documentName As String) As String
' Returns cleaned GSC document type from inputted document name - needs to be opened - or null

On Error GoTo gdError

Dim tmpDocType As String

If Is_GSC_Doc(documentName) Then
    
    tmpDocType = CleanGSC_DocType(Left$(documentName, 2))
    
    If tmpDocType <> "" Then
        getDocType_fromDocName = tmpDocType
        Exit Function
    Else
        getDocType_fromDocName = ""
        Exit Function
    End If
    
Else
    getDocType_fromDocName = ""
    Exit Function
End If

Exit Function

'___________________________________

gdError:
    If Err.Number <> 0 Then
        Err.Clear
        getDocType_fromDocName = ""
    End If

End Function


Function getLanguage_fromDocName(documentName As String) As String
' Returns valid GSC document language from inputted document name - needs to be opened - or null

On Error GoTo glError

Dim tmpLang As String

If Is_GSC_Doc(documentName) Then
    
    tmpLang = Left$(Split(documentName, ".")(1), 2)
    
    If gscrolib.IsValidLangIso(tmpLang) Then
        getLanguage_fromDocName = tmpLang
        Exit Function
    Else
        getLanguage_fromDocName = ""
        Exit Function
    End If
    
Else
    getLanguage_fromDocName = ""
    Exit Function
End If

Exit Function

'___________________________________

glError:
    If Err.Number <> 0 Then
        Err.Clear
        getLanguage_fromDocName = ""
    End If

End Function


Function Build_MProd_Path(HumanReadableGSC_DocumentID As String, Optional getFileNameOnly) As String
' returns empty string "" if not found
' SN 1125/11 REV1 or 15221/11 or 15221/13 AD1RE3

Dim dtype As String
Dim dnum As String
Dim dyea As String
Dim dsuf As String

Dim snum As Variant     ' suffix numbers, extraction string array

'Dim getFileNameOnly As String
'getFileNameOnly = "False"

Dim idoc As String
idoc = Trim(HumanReadableGSC_DocumentID)

If UBound(Split(Split(idoc, "/")(0), " ")) = 0 Then     ' No type given in input string
    dtype = "st"
    dnum = Format(Split(idoc, "/")(0), "00000")
Else
    dtype = LCase(Split(Split(idoc, "/")(0), " ")(0))
    dnum = Format(Split(Split(idoc, "/")(0), " ")(1), "00000")
End If

' Extract suffix and year
If UBound(Split(Split(idoc, "/")(1), " ")) = 0 Then     ' No suffix in given input string
    dsuf = ""
    dyea = Trim(Split(idoc, "/")(1))
Else

    dsuf = (Replace(Split(Split(idoc, "/")(1), " ")(1), " ", ""))
    If Is_StdSuff(dsuf) Then
        dsuf = "-" & SGC_StdSuff_To_FileSuff(dsuf)
    Else
        If Is_WFSuff(dsuf) Then
           dsuf = "-" & SGC_StdSuff_To_FileSuff(dsuf)
        Else
            ' we found a space in the second part of the string, split after /, but
            ' the remaining string is NOT a standard suffix string. What DO?
            Build_MProd_Path = ""
            Exit Function
        End If
    End If

    dyea = Split(Split(idoc, "/")(1), " ")(0)
    'dsuf = Replace(Right$(dsuf, 1))
End If

Dim mProdbasePath As String
mProdbasePath = MProdBaseFolder & dyea & "\" & dtype & Left(dnum, 2)

If Dir(mProdbasePath & "\" & dtype & dnum & dsuf & ".??" & dyea & ".docx") <> "" Then
    If Not IsMissing(getFileNameOnly) Then
        If getFileNameOnly = "True" Or (LCase(getFileNameOnly) = "getfilenameonly") Then
            Build_MProd_Path = Dir(mProdbasePath & "\" & dtype & dnum & dsuf & ".??" & dyea & ".docx")
        Else
            Build_MProd_Path = mProdbasePath & "\" & Dir(mProdbasePath & "\" & dtype & dnum & dsuf & ".??" & dyea & ".docx")
        End If
    Else
        Build_MProd_Path = mProdbasePath & "\" & Dir(mProdbasePath & "\" & dtype & dnum & dsuf & ".??" & dyea & ".docx")
    End If
Else
    If Dir(mProdbasePath & "\" & dtype & dnum & dsuf & ".??" & dyea & ".txt") <> "" Then
        If Not IsMissing(getFileNameOnly) Then
            If getFileNameOnly = "True" Or (LCase(getFileNameOnly) = "getfilenameonly") Then
                Build_MProd_Path = Dir(mProdbasePath & "\" & dtype & dnum & dsuf & ".??" & dyea & ".txt")
            Else
                Build_MProd_Path = mProdbasePath & "\" & Dir(mProdbasePath & "\" & dtype & dnum & dsuf & ".??" & dyea & ".txt")
            End If
        Else
            Build_MProd_Path = mProdbasePath & "\" & Dir(mProdbasePath & "\" & dtype & dnum & dsuf & ".??" & dyea & ".txt")
        End If
    Else
        Build_MProd_Path = ""
    End If
End If

End Function

Function Found_Opened_Document(DocFilename As String) As Boolean
' v 0.2

If Documents.count = 0 Then
    Found_Opened_Document = False
    Exit Function
End If

Dim doc As Document

For Each doc In Documents
    If getGSCBaseName(doc.Name) = getGSCBaseName(DocFilename) Then
        Found_Opened_Document = True
        Exit Function
    End If
Next doc

End Function


Function GetEnvironmentVar(Name As String) As String

GetEnvironmentVar = String(255, 0)
GetEnvironmentVariable Name, GetEnvironmentVar, Len(GetEnvironmentVar)
GetEnvironmentVar = TrimNull(GetEnvironmentVar)

End Function

Private Function TrimNull(item As String)

Dim iPos As Long
iPos = InStr(item, vbNullChar)
TrimNull = IIf(iPos > 0, Left$(item, iPos - 1), item)

End Function

Sub AutoBackupLocation_Initialize()
' Sets reg key "HKEY_CURRENT_USER\Software\Microsoft\Office\10.0\Word\GSCRo Settings\AutoBackupLocation"

Dim fso As New Scripting.FileSystemObject

If AutoBackupLocation = "" Then
    If gscrolib.NoRegKey("GSCRo Settings", "AutoBackupLocation") Then    ' If regkey does not exist
        System.ProfileString("GSCRo Settings", "AutoBackupLocation") = _
            GetDirectory("Please select folder for AutoBackup Location")
            AutoBackupLocation = System.ProfileString("GSCRo Settings", "AutoBackupLocation")
    Else    ' If key exists
        AutoBackupLocation = System.ProfileString("GSCRo Settings", "AutoBackupLocation")
    End If
End If

Set fso = Nothing
End Sub

Function IsValidLangIso(InputString As String) As Boolean
' Boolean function, returns whether given input string is a valid language iso as defined at GSC

Dim ValidLgIso(24) As String
Dim i As Integer

ValidLgIso(0) = "en": ValidLgIso(12) = "cs"
ValidLgIso(1) = "fr": ValidLgIso(13) = "et"
ValidLgIso(2) = "ro": ValidLgIso(14) = "lv"
ValidLgIso(3) = "de": ValidLgIso(15) = "lt"
ValidLgIso(4) = "nl": ValidLgIso(16) = "hu"
ValidLgIso(5) = "it": ValidLgIso(17) = "mt"
ValidLgIso(6) = "es": ValidLgIso(18) = "pl"
ValidLgIso(7) = "da": ValidLgIso(19) = "sk"
ValidLgIso(8) = "el": ValidLgIso(20) = "sl"
ValidLgIso(9) = "pt": ValidLgIso(21) = "bg"
ValidLgIso(10) = "fi": ValidLgIso(22) = "xx"
ValidLgIso(11) = "sv": ValidLgIso(23) = "hr"
ValidLgIso(24) = "ga"     ' Rarer gaelic - irish

For i = 0 To UBound(ValidLgIso)
    If LCase(InputString) = ValidLgIso(i) Then
        IsValidLangIso = True
        Exit Function
    End If
Next i


End Function

Function getGSCBaseName(documentName As String) As String

Dim fso As New Scripting.FileSystemObject
Dim tmpDocName As String

If Not Is_GSC_Doc(documentName) Then
    getGSCBaseName = ""
    Exit Function
End If

Dim dotIndex As Integer
dotIndex = InStr(1, documentName, ".")

Dim baseLength As Integer
baseLength = dotIndex + 4

getGSCBaseName = Left$(documentName, baseLength)

End Function

Sub MakeForm_Resizable(TargetUserForm As UserForm)
    
    Dim UFHWnd As Long    ' HWnd of UserForm
    Dim WinInfo As Long   ' Values associated with the UserForm window
    Dim r As Long
    'Const GWL_STYLE = -16
    
    Const WS_SIZEBOX = &H40000
    Const WS_USER  As Long = &H4000
    Load TargetUserForm ' Load the form into memory but don't make it visible
    
    UFHWnd = FindWindow("ThunderDFrame", TargetUserForm.Caption)  ' find the HWnd of the UserForm
    If UFHWnd = 0 Then
        ' cannot find form
        Debug.Print "Userform " & TargetUserForm.Name & "not found"
        Exit Sub
    End If
    
    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)      ' get the style word (32-bit Long)
    WinInfo = WinInfo Or WS_SIZEBOX                 ' set the WS_SIZEBOX bit
    r = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)   ' set the style word to the modified value.
    
End Sub

Sub AhMsgBox_Show(TextToShow As String, Optional CountDownTimer As Integer, Optional PositionCounter, _
                    Optional NewWidth As Integer, Optional NewHeight As Integer, _
                    Optional NewBackColor As WdColor, Optional NewForeColor As WdColor)

' version 0.2 - now can use optional parameter to decide where to print (to stack multiple messages)

'Dim nUser As New FormHndTest

With FormHndTest
    
    .Label1.Caption = TextToShow
    .BorderStyle = fmBorderStyleNone
    .Label1.BackStyle = fmBackStyleTransparent
        
    If NewBackColor <> 0 Then
        .BackColor = NewBackColor
        .BorderColor = NewBackColor
    Else
        .BackColor = wdColorLightYellow
        .BorderColor = wdColorLightYellow
    End If
    
    If NewForeColor <> 0 Then
        .Label1.ForeColor = NewForeColor
        .Label2.ForeColor = NewForeColor
    Else
        .Label1.ForeColor = wdColorIndigo
        .Label2.ForeColor = wdColorIndigo
    End If
    
    If CountDownTimer <> 0 Then
        FormHndTest.NewCountDownTime = CountDownTimer
    End If
    
    If NewWidth <> 0 Then
        .Width = NewWidth
    End If
    
    If NewHeight <> 0 Then
        .Height = NewHeight
    End If
    
    With .Label1.Font
        .Name = "Trebuchet MS"
        .Size = 10
        .Bold = True
    End With
    
    If Not IsMissing(PositionCounter) Then
        nUser.nthAhMsg = PositionCounter
    End If
    
End With

FormHndTest.Show

End Sub

Function IsVisible_CommandBar(CommandBarName As String) As Boolean
' Returns True if commandbar is found and visible, False otherwise

Dim tcb As CommandBar

For Each tcb In Application.CommandBars
    If tcb.Name = CommandBarName Then
        IsVisible_CommandBar = True
        Exit Function
    End If
Next tcb

IsVisible_CommandBar = False

End Function

Function IsVisibleUserForm(UserFormName As String) As Boolean
' Returns True if userform is found and visible, False otherwise

Dim OD_hwnd As Long

OD_hwnd = gscrolib.FindWindow("ThunderDFrame", "OpenDoc")

If Not OD_hwnd = 0 Then
    IsVisibleUserForm = True
Else
    IsVisibleUserForm = False
End If

End Function

Sub SetMousePos_C(XCoord As Long, YCoord As Long)
' Move mouse cursor at specified position, mediated so that it would not appear near the edge

Dim dl As Long
Dim pt As POINTAPI

dl = GetCursorPos(pt)

If dl <> 0 Then
    
    If pt.y <= 450 Then
        pt.x = XCoord
        pt.y = pt.y + YCoord
    Else
        pt.x = 200
        pt.y = pt.y - YCoord
    End If
    
    SetCursorPos pt.x, pt.y
End If

End Sub

Sub SetMousePosition(XCoord As Long, YCoord As Long)
' Move mouse cursor at specified position, mediated so that it would not appear near the edge

Dim dl As Long      ' Result of GetCursor operation - non zero = success, 0 = failure
Dim pt As POINTAPI  ' Coordinates of mouse cursor arrive in this structure

'dl = GetCursorPos(pt)

'If dl <> 0 Then
    SetCursorPos pt.x, pt.y
'End If

End Sub


Function NoRegKey(SubSection As String, KeyName As String) As Boolean

On Error GoTo NoSubSect

Select Case System.ProfileString(SubSection, KeyName)

    Case Else
        NoRegKey = False
End Select


NoSubSect:
If Err = 5843 Then
    Err.Clear
    NoRegKey = True
End If

End Function
Function GetDirectory(Optional msg) As String
    On Error Resume Next
    Dim bInfo As BrowseInfo
    Dim path As String
    Dim r As Long, x As Long, pos As Integer
     
     'Root folder = Desktop
    bInfo.pIDLRoot = 0&
     
     'Title in the dialog
    If IsMissing(msg) Then
        bInfo.lpszTitle = "Please select a folder"
    Else
        bInfo.lpszTitle = msg
    End If
     
     'Type of directory to return
    bInfo.ulFlags = &H1 Or &H40     ' &H40 is needed to display new version of browse window, allows for folder creation
                                    ' and has shortcut menu for folder/ files
     
     'Display the dialog
    x = SHBrowseForFolder(bInfo)
     
     'Parse the result
    path = Space$(512)
    r = SHGetPathFromIDList(ByVal x, ByVal path)
    If r Then
        pos = InStr(path, Chr$(0))
        GetDirectory = Left(path, pos - 1)
    Else
        GetDirectory = ""
    End If
End Function

Function ListAllSubfolders(RootFolder As String, Optional GetFilesAlso As Boolean) As Variant()
' version 0.1

Dim lasFo As New Scripting.FileSystemObject

Dim lafRA()

racounter = 0

If lasFo.FolderExists(RootFolder) = False Then
    MsgBox "Folderul " & RootFolder & " nu exista!" & vbCr & "Rugam verificati si relansati!", vbOKOnly + vbCritical, "No such thing!"
    End
End If

Call GetAllSubfolders(RootFolder, lafRA())
ListAllSubfolders = lafRA()

Debug.Print ""

End Function
Sub GetAllSubfolders(GARootFolderS As String, Optional resultArray)
' version 0.1
' ATTENTION, recursive subroutine !

Dim GafFo As New Scripting.FileSystemObject
Dim GARootFolder As Folder
Dim sf As Folder

Set GARootFolder = GafFo.GetFolder(GARootFolderS)

If GARootFolder.SubFolders.count > 0 Then
    For Each sf In GARootFolder.SubFolders
        racounter = racounter + 1
        ReDim Preserve resultArray(racounter - 1)
        resultArray(racounter - 1) = sf.path
        Call GetAllSubfolders(sf.path, resultArray)
    Next sf
End If

End Sub

Sub Subfolders_Test()

Dim v As Variant

v = ListAllSubfolders("D:\VBA")

End Sub


Sub msgbox_pausedisplay(message As String, userformcaption As String, pauseinseconds As Single, _
                        Optional NewWidth As Integer, Optional NewHeight As Integer)

frm_msgbox_pause.Caption = userformcaption
frm_msgbox_pause.Label1.Caption = message
frm_msgbox_pause.Show vbModeless

If NewWidth <> 0 Then
    
    With frm_msgbox_pause
        .Width = NewWidth
        .Label1.Width = .Width - 30
        .CommandButton1.Left = .Width / 2 - 21
    End With
    
End If

If NewHeight <> 0 Then
    frm_msgbox_pause.Height = NewWidth
End If

Call PauseForSeconds(pauseinseconds)
Unload frm_msgbox_pause

End Sub


Sub PauseForSeconds(numberofseconds As Single, Optional Bailout)

Dim PauseTime, StartTime, FinishTime, TotalTime

PauseTime = numberofseconds             ' Set duration.
StartTime = Timer                       ' Set start time.

Do While Timer < StartTime + PauseTime
    
    If Not IsMissing(Bailout) Then
        Exit Sub
    End If
    
    DoEvents                            ' Yield to other processes.
Loop

FinishTime = Timer                      ' Set end time.
TotalTime = FinishTime - StartTime      ' Calculate total time.
    

End Sub


Sub ArrangeSideBySideW()

If NoDoc Then Exit Sub
If Not WToff Then Exit Sub

If Application.Documents.count < 2 Then
    asmsg = "Please open TWO documents before running this macro." & vbCr & vbCr & _
         "(The document you are in will be on the right-hand side of the screen)"
    Call gscrolib.msgbox_pausedisplay(CStr(asmsg), "Scroll Impossible", 3)
    Exit Sub
End If

ActiveWindow.WindowState = wdWindowStateNormal
    With ActiveWindow
        .Left = .UsableWidth / 2
        .Top = 0
        .Width = .UsableWidth / 2
        .Height = .UsableHeight
    End With
ActiveWindow.ActivePane.View.Type = wdPrintView
ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitFullPage
    
On Error GoTo Err
    
ActiveDocument.ActiveWindow.Next.Activate
    With ActiveWindow
        .Width = .UsableWidth / 2
        .Height = .UsableHeight
        .Left = 0
        .Top = 0
    End With
ActiveWindow.ActivePane.View.Type = wdPrintView
ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitFullPage
    
Exit Sub
    
Err:
Select Case Err.Number
    Case 91
        ActiveDocument.ActiveWindow.Previous.Activate
        Resume Next
    Case Else
        as1msg = Err.Number & "  " & Err.Description
        Call gscrolib.msgbox_pausedisplay(CStr(as1msg), "Scroll Error", 3)
End Select
    
End Sub
Public Function GetListSeparatorFromRegistry() As String

lnErrnum = 0
        
Dim sLInit As String
Dim sL As String
        
sLInit = System.PrivateProfileString(vbNullString, "HKEY_CURRENT_USER\Control Panel\International", "sList")
sL = sLInit
    
If Documents.count > 0 Then
    'Test validity of sL and trap error msg
    With Selection.Find
        .Text = " {2" & sL & "}"
        .Forward = True
        .Format = False
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .MatchAllWordForms = False
    End With

Dim fSelection As Range
Set fSelection = Selection.Range    ' save user selection for later restore... sooo ruude !

End If


On Error Resume Next

Selection.Find.Execute
lnErrnum = Err.Number
Err.Clear

fSelection.Select

If lnErrnum = 5560 Then 'Wrong list separator
    If sL = ";" Then 'Change
        sL = ","
        GetListSeparatorFromRegistry = sL
    Else
        If sL = "," Then 'Change
            sL = ";"
            GetListSeparatorFromRegistry = sL
        End If
    End If
Else
    GetListSeparatorFromRegistry = sL
End If
Call ClearFindReplace
    
End Function
Public Sub SetListSeparator(ByRef strSeparator As String)
    System.PrivateProfileString("", "HKEY_CURRENT_USER\Control Panel\International", "sList") = strSeparator
End Sub
Public Sub ClearFindReplace()
    '---------------------------------------------------------------------------------------
    ' Procedure : ClearFindReplace
    ' DateTime  : 9/08/2006 08:39
    ' Author    : KUHNRAY
    ' Purpose   : Enlève texte(s) ou paragraphe(s) formaté(s)
    '---------------------------------------------------------------------------------------
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
    End With
    
    FndRpl "", "", True, 1, False, False, False, False, False, False
End Sub
Function FndRpl(TxtToRpl, RplTxt, _
                    Optional Forward As Boolean, _
                    Optional Wrap As Integer, _
                    Optional Format As Boolean, _
                    Optional MCase As Boolean, _
                    Optional MWholeWord As Boolean, _
                    Optional MWildcards As Boolean, _
                    Optional MSoundsLike As Boolean, _
                    Optional MAWForms As Boolean)

With Selection.Find
    .Text = TxtToRpl
    .Replacement.Text = RplTxt
    .Forward = Forward
    .Wrap = Wrap
    .Format = Format
    .MatchCase = MCase
    .MatchWholeWord = MWholeWord
    .MatchWildcards = MWildcards
    .MatchSoundsLike = MSoundsLike
    .MatchAllWordForms = MAWForms
End With

Selection.Find.Execute Replace:=wdReplaceAll

End Function
Sub TrackChanges_Off(documentName As String)
' version 0.1
' main routine
'
' could be modified into a function, to return whether the trackchanges were on or off, or/and to
' accept a parameter to accept all trackchanges or reject all before turning them off.

Dim cdoc As Document

Set cdoc = Documents.Open(FileName:=documentName, Visible:=False, AddToRecentFiles:=False)

If cdoc.TrackRevisions = True Then
    If cdoc.Revisions.count > 0 Then
        For Each myrev In cdoc.Revisions
            myrev.Accept
        Next myrev
    End If
    cdoc.TrackRevisions = False
    Application.DisplayAlerts = wdAlertsNone
    cdoc.Close SaveChanges:=wdSaveChanges
Else
    cdoc.Close SaveChanges:=wdDoNotSaveChanges
End If
    
End Sub
Function IsNotDimensionedArr(arrayname, Optional whichdimension) As Boolean
' version 0.4
' Functia returneaza True doar daca matricea (array-ul) furnizat ca parametru este ne-dimensionat.
' Returneaza True de asemenea, daca userul cere info (furnizat de j) despre o dimensiune inexistenta
' a matricii. Va returna False in cazul in care matricea este dimensionata deja, dar are toate
' elementele vide, nule, goale, etc.


On Error GoTo handleme
If IsMissing(whichdimension) Then
    If UBound(arrayname) >= 0 Then
        IsNotDimensionedArr = False
    End If
Else
    If UBound(arrayname, whichdimension) >= 0 Then
        IsNotDimensionedArr = False
    End If
End If

GoTo EndMe
handleme:
If Err.Number = 9 Then      ' Eroarea 9 inseamna "Subscript out of range" si este
    Err.Clear
    IsNotDimensionedArr = True
End If

EndMe:
End Function


Function Get_ULg_Old() As String
' version 3.0   ' accounting for IOLAN, VDI, ROLAN (new), Classic TW
' Function returns empty if computer name not GSC standard, such as "PLRO..." or "TRRO..." or TW computer (clientname env var used)



Dim tmpLang As String
tmpLang = sComputerName

' Only works for GSC IOLAN computers...
If Left$(tmpLang, 2) = "PL" Or Left$(tmpLang, 2) = "TR" Then
    
    If IsValidLangIso(Mid$(tmpLang, 3, 2)) Then
    
        Get_ULg = LCase(Mid$(tmpLang, 3, 2))
        Exit Function
        
    End If
    
End If


' Classic TW case
If Left$(tmpLang, 2) = "CS" Then
    
    tmpLang = Mid$(GetEnvironmentVar("clientname"), 3, 2)
    
    If IsValidLangIso(tmpLang) Then
    
        Get_ULg = LCase(tmpLang)
        Exit Function
        
    End If
    
End If


' ROLAN (new version) case
If GetEnvironmentVar("clientname") <> "" Then
    
    tmpLang = Left$(GetEnvironmentVar("clientname"), 2)
    
    If IsValidLangIso(tmpLang) Then
    
        Get_ULg = LCase(tmpLang)
        Exit Function
        
    End If
    
End If


' DVI case
If GetEnvironmentVar("OU") <> "" Then
    
    tmpLang = Replace(Split(GetEnvironmentVar("OU"), "&")(1), "TRAD", "")
    
    If IsValidLangIso(tmpLang) Then
    
        Get_ULg = LCase(tmpLang)
        Exit Function
        
    End If
    
End If

' Function will return empty string if no case above worked

End Function


Function Is_GSC_Doc(FileName As String, Optional SilentMode) As Boolean

' version 0.7
' independent
' depends on VerificaCN
' 0.5 - Added XM, XN, XT
' 0.6 - Minding only first 2 groups of chars separated by dot.
' 0.7 - Adding some missing doc types (EG, )

Dim MsgS As String
Dim PointPos As Byte
Dim LinePos As Byte
Dim ff As New Scripting.FileSystemObject

' Is filename extensioned or not ? (".doc" extension)
If Right$(FileName, 5) = ".docx" Or Right$(FileName, 4) = ".doc" Then
    ' This way we ONLY look at
    If UBound(Split(FileName, ".")) = 2 Then
        BaseDocName = ff.GetBaseName(FileName)
    Else
        BaseDocName = Split(FileName, ".")(0) & "." & Split(FileName, ".")(1)
    End If
Else
    BaseDocName = FileName
End If

PointPos = InStr(1, BaseDocName, ".")
    
Select Case PointPos
    Case 0
        MsgS = "Not Council document, or not properly named! Cherish the point!"

Ooops:  If IsMissing(SilentMode) Then
            StatusBar = MsgS
        End If
        
        Is_GSC_Doc = False
        Exit Function
        
    Case Is = 8
        LinePos = InStr(1, BaseDocName, "-", vbTextCompare)
            If LinePos <> 0 And LinePos <= PointPos Then MsgS = _
            "Not Council document, or not properly named! Where's your line at?": GoTo Ooops
    Case Is > 9
        If (PointPos - 9) Mod 4 <> 0 Then MsgS = _
        "Not Council document, or not properly named! Where's your point at?": GoTo Ooops
        LinePos = InStr(1, BaseDocName, "-", vbTextCompare)
        If LinePos <> 8 Then MsgS = _
        "Not Council document, or not properly named! Cherish your lines!": GoTo Ooops
    Case Else
        MsgS = "Not Council document, or not properly named! Where's your point at?": GoTo Ooops
End Select

Set ff = Nothing

DocType(0, 0) = "st": DocType(0, 7) = "cp"
DocType(0, 1) = "sn": DocType(0, 8) = "ad"
DocType(0, 2) = "cm": DocType(0, 9) = "ac"
DocType(0, 3) = "ds": DocType(0, 10) = "nc"
DocType(0, 4) = "lt": DocType(0, 11) = "np"
DocType(0, 5) = "bu": DocType(0, 12) = "da"
DocType(0, 6) = "pe": DocType(0, 13) = "rs"
DocType(0, 14) = "cg": DocType(0, 15) = "fa"
DocType(0, 16) = "xm": DocType(0, 17) = "xn"
DocType(0, 18) = "xt": DocType(0, 19) = "eg"
DocType(0, 20) = "re": DocType(0, 21) = "de"


DocType(1, 0) = "1"
DocType(1, 1) = "2"

DocSuff(0, 0) = "re": DocSuff(0, 3) = "ex"
DocSuff(0, 1) = "co": DocSuff(0, 4) = "am"
DocSuff(0, 2) = "ad": DocSuff(0, 5) = "dc"

DocSuff(1, 0) = 8
DocSuff(1, 1) = PointPos - 8

ReDim DocYear(1, (Right$(DatePart("yyyy", Date), 2)) - 1) 'Ultimele doua cifre ale anului curent
j = -1

For i = Right$(DatePart("yyyy", Date), 2) To 1 Step -1
    j = j + 1
    DocYear(0, j) = Format(i, "00")
Next i

DocYear(1, 0) = PointPos + 3
DocYear(1, 1) = "2"

EUDocLang(0, 0) = "en": EUDocLang(0, 12) = "cs"
EUDocLang(0, 1) = "fr": EUDocLang(0, 13) = "et"
EUDocLang(0, 2) = "ro": EUDocLang(0, 14) = "lv"
EUDocLang(0, 3) = "de": EUDocLang(0, 15) = "lt"
EUDocLang(0, 4) = "nl": EUDocLang(0, 16) = "hu"
EUDocLang(0, 5) = "it": EUDocLang(0, 17) = "mt"
EUDocLang(0, 6) = "es": EUDocLang(0, 18) = "pl"
EUDocLang(0, 7) = "da": EUDocLang(0, 19) = "sk"
EUDocLang(0, 8) = "el": EUDocLang(0, 20) = "sl"
EUDocLang(0, 9) = "pt": EUDocLang(0, 21) = "bg"
EUDocLang(0, 10) = "fi": EUDocLang(0, 22) = "xx"
EUDocLang(0, 11) = "sv": EUDocLang(0, 23) = "hr"
EUDocLang(0, 24) = "ga"     ' Rarer gaelic - irish

EUDocLang(1, 0) = PointPos + 1
EUDocLang(1, 1) = 2


For i = 1 To 5
    If VerificaCN(i, BaseDocName) = False Then
        MsgS = "Not Council Document, Sorry!"
        GoTo Ooops
    End If
Next i

Is_GSC_Doc = True

End Function

Function VerificaCN(VerifN, VString As String) As Boolean   ' Verifica Council Name
' version 1.0
' independent
' Added "DCL" suffix to represent declasiffications

Dim CeVerificam
Dim EValid As Boolean, ENumeric As Boolean

Select Case VerifN
Case Is = 1
        CeVerificam = DocType()
VLoop:  For j = 0 To UBound(CeVerificam, 2)
            If StrComp(Mid$(BaseDocName, CeVerificam(1, 0), CeVerificam(1, 1)), _
            CeVerificam(0, j), vbTextCompare) = 0 Then
                VerificaCN = True
                Exit For
            End If
        Next j

Case Is = 2
    If IsNumeric(Mid$(BaseDocName, 3, 5)) = True Then VerificaCN = True
        
Case Is = 3
    If InStr(1, BaseDocName, "-", vbTextCompare) = 0 Then
        VerificaCN = True
    ElseIf InStr(1, BaseDocName, "-", vbTextCompare) = 8 Then
        For n = 1 To (DocSuff(1, 1) - 1) / 4
            EValid = False
            For j = 0 To UBound(DocSuff, 2)
                If StrComp(Mid$(BaseDocName, (8 + (4 * n - 3)), 2), DocSuff(0, j), vbTextCompare) = 0 Then
                    EValid = True
                    Exit For
                End If
            Next j
            If EValid = False Then Exit For
        Next n
        For n = 1 To (DocSuff(1, 1) - 1) / 4
        ENumeric = False
            If IsNumeric(Mid$(BaseDocName, (8 + (4 * n - 1)), 2)) = True Then ENumeric = True
            If ENumeric = False Then Exit For
        Next n
        
        If EValid = True And ENumeric = True Then VerificaCN = True
    End If
Case Is = 4
    CeVerificam = EUDocLang()
    GoTo VLoop
Case Is = 5
    CeVerificam = DocYear()
    GoTo VLoop
End Select

End Function

Sub Set_currWUser()
' version 0.8
' secondary to Open_finalDoc
'
'

Dim vfe As New Scripting.FileSystemObject

If Application.UserName <> "" Then
    If vfe.FolderExists("T:\Docs\" & Application.UserName) Then
        currWUser = Application.UserName
    Else
        Application.UserName = InputBox("Va rugam introduceti user-name-ul dvs de windows.", _
        "Username nu e setat in Word!", Application.UserName)
        If vfe.FolderExists("T:\Docs\" & Application.UserName) Then
            currWUser = Application.UserName
        End If
    End If
End If

Set vfe = Nothing
End Sub

Function GSCFileName_ToStandardName(FileName As String, Optional SpacedSuffix, Optional NoPrefixforST, Optional RevNumInDocNum, Optional UseWFSuffix) As String
' version 0.5
' independent
'
' Convert string representing gsc document filename to standard name (st12345.en12.doc to ST 12345/12)
' Added option to produce standard name with REV number before year, as in: st12345-re01.en13.doc -> ST 12345/1/13 REV1
' RevNumInDocNum can be set to
'    0 or "RevNumInSuffix", so as to NOT include rev number in doc number, list it as suffix, as in ST 12345/13 REV1
' or 1, or "RevNumInDocNum", so as to include it in doc number AND list it as suffix as well, example ST 12345/1/13 REV1
' or 2, or "RevNumInDocNumOnly", so as to include it ONLY in document number, not in suffix, example ST 12345/1/13
' 0.5 - Use Workflow suffix - Added optional parameter to use spaceless suffixes and contract suffixes to2 letters + 0-less numbers, as in WF!


Dim SType As String, SNumber As String, SYear As String, SSuff As String
Dim SPointPos As Byte
SPointPos = InStr(1, FileName, ".")

SType = UCase(Left$(FileName, 2))
SNumber = Format(Mid$(FileName, 3, 5), "#####")
SYear = Mid$(FileName, SPointPos + 3, 2)

If SPointPos > 8 Then

    If InStr(1, FileName, "-") > 0 Then

        If Not IsMissing(SpacedSuffix) Then
            If CStr(SpacedSuffix) = "True" Or LCase(CStr(SpacedSuffix)) = "spacedsuffix" Then
                SSuff = Mid$(FileName, 9, (SPointPos - 9))
                SSuff = SGC_Suff_ToStandard_Suff(SSuff)
            ElseIf CStr(SpacedSuffix) = "False" Or LCase(CStr(SpacedSuffix)) = "nospacedsuffix" Then
                SSuff = Mid$(FileName, 9, (SPointPos - 9))
                SSuff = SGC_Suff_ToStandard_SuffNS(SSuff)
            Else    ' If spacedsuffix is wrongly specified, default is "no space" version
                SSuff = Mid$(FileName, 9, (SPointPos - 9))
                SSuff = SGC_Suff_ToStandard_SuffNS(SSuff)
            End If
        Else
            ' Non spaced suffix by default
            SSuff = Mid$(FileName, 9, (SPointPos - 9))
            SSuff = SGC_Suff_ToStandard_SuffNS(SSuff)

            ' contract to WF suffixes, as requested
            If Not IsMissing(UseWFSuffix) Then

                SSuff = SGC_StandardSuff_toWFSuff(SSuff)

            End If


        End If
    End If
End If

' Handle revision number in doc number inclusion or not
If Not IsMissing(RevNumInDocNum) Then
    Select Case LCase(CStr(RevNumInDocNum))
        Case "0", "revnuminsuffix"
            ' No need to do, default behavious anyway
        Case "1", "revnumindocnum"
            If SSuff <> "" Then
                SNumber = SNumber & "/" & IIf(InStr(1, SSuff, "REV"), Extr_Numbers(SSuff)(0), SNumber)
            End If
        Case "2", "revnumindocnumonly"
            If SSuff <> "" Then
                SNumber = SNumber & "/" & IIf(InStr(1, SSuff, "REV"), Extr_Numbers(SSuff)(0), SNumber)
                ' Now, that's an expression, ha ? Just had to do it !
                SSuff = Trim(Replace(SSuff, Left$(SSuff, (IIf(Mid$(SSuff, 4, 1) = " ", 5, 4))), ""))
            End If
        Case Else
    End Select
End If

' we assemble the final string, with or without suffix
If SSuff <> "" Then
    GSCFileName_ToStandardName = SType & " " & SNumber & "/" & SYear & " " & UCase(SSuff)
Else
    GSCFileName_ToStandardName = SType & " " & SNumber & "/" & SYear
End If

' We correct prefix (document type) as per user request (discard it for ST docs)
If Not IsMissing(NoPrefixforST) Then
    ' Any value will do, such as the name of parameter itself, or true or 1
    If SType = "ST" Then
        GSCFileName_ToStandardName = Replace(GSCFileName_ToStandardName, "ST ", "")
    Else
    End If

Else    ' If parameter is missing, default behaviour is to leave prefix in
End If

End Function


Function SGC_StandardSuff_toWFSuff(suffix As String) As String

Dim tmpWFSuff As String

tmpWFSuff = Replace(suffix, " ", "")

tmpWFSuff = Replace(tmpWFSuff, "REV", "RE")
tmpWFSuff = Replace(tmpWFSuff, "ADD", "AD")
tmpWFSuff = Replace(tmpWFSuff, "COR", "CO")
tmpWFSuff = Replace(tmpWFSuff, "AMD", "AM")
tmpWFSuff = Replace(tmpWFSuff, "EXT", "EX")
tmpWFSuff = Replace(tmpWFSuff, "DCL", "DC")

tmpWFSuff = Replace(tmpWFSuff, "0", "")

SGC_StandardSuff_toWFSuff = tmpWFSuff

End Function


Function GSC_CarsNameFormat_ToStandardFormat(CARSNameFormat As String, Optional IgnoreSuppliedLanguage) As String

' Just transforming 1515 2018 into 1515/18, leaving suffix, doc type alone

' v 0.2 - Addded optional param IgnoreSuppliedLanguage so as to eliminate supplied language code from result (useful for multiple projects backup check program)

Dim suppliedString As String
suppliedString = Replace(CARSNameFormat, Chr(160), " ")

Dim inputWordsArr
inputWordsArr = Split(suppliedString, " ")


If UBound(inputWordsArr) < 1 Or _
    UBound(inputWordsArr) > 6 Then      ' Accepting 3 suffixes, maximum
    
    GSC_CarsNameFormat_ToStandardFormat = ""
    Exit Function
    
End If


If CleanGSC_DocType(CStr(inputWordsArr(0))) <> "" Then
    
    ' cant be smaller than 3 elems in array, we need type number and year to be valid !
    If UBound(inputWordsArr) < 2 Then
        GSC_CarsNameFormat_ToStandardFormat = ""
        Exit Function
    End If
    
    If IsNumeric(inputWordsArr(1)) And IsNumeric(inputWordsArr(2)) Then
        inputWordsArr(2) = "/" & Right(inputWordsArr(2), 2)  ' just shorten year
    Else
        GSC_CarsNameFormat_ToStandardFormat = ""
        Exit Function
    End If
    
Else    ' first elem is not doc type, we have an ST !
        
    If UBound(inputWordsArr) < 1 Then   ' thats acceptable minimum, we need at least doc number and year !
        GSC_CarsNameFormat_ToStandardFormat = ""
        Exit Function
    End If
    
    If IsNumeric(inputWordsArr(0)) And IsNumeric(inputWordsArr(1)) Then
        inputWordsArr(1) = "/" & Right(inputWordsArr(1), 2)  ' just shorten year
    Else
        GSC_CarsNameFormat_ToStandardFormat = ""
        Exit Function
    End If
    
End If


Dim tmpResult As String
tmpResult = Join(inputWordsArr, " ")

' return result with or without language code
If Not IsMissing(IgnoreSuppliedLanguage) Then
    If IsValidLangIso(Right(tmpResult, 2)) Then
        tmpResult = Trim(Left(tmpResult, Len(tmpResult) - 2))
    End If
End If

tmpResult = Trim(Replace(tmpResult, "INIT", ""))

GSC_CarsNameFormat_ToStandardFormat = Replace(tmpResult, " /", "/")
    

End Function

Function GSC_StandardName_ToFileName(StandardName As String, Optional DocLanguage, Optional WithFileExtension) As String
' version 0.86

' Converts a standard document name to GSC filename as in "ST 12345/12 ADD1 REV2" to "st12345-ad01re02.en12"
' 0.70: corrected grave error in which it did not know to handle doc id such as 12536/1/12 (to result in st12536-re01.xx12),
' because it did not recognize that revision number may be embedded in the number-year string, or type-number-year string
' 0.71: made it not fail if supplied with arbitrary string
' 0.72: add poss to
' 0.8: added handling of short compounded suffixes, Workflow style: "RE1CO1" is now legal!
' 0.81: function not recognising "AD 16/2016", but same works for "SN" or "ST" !
' 0.84: support new doc ID format, namely Automate & CARS Search client format: SN 2525 2018 REV1
' 0.86: fixed bug (error whith certain non-conforming strings)

Dim SType As String, SNumber As String, SYear As String, SSuff As String, SLng As String
Dim StdNameArr() As String
Dim ERevisionNum As String      ' Embedded revision number

Dim suppliedStandardName As String

' Sanity check
If InStr(1, StandardName, "/") = 0 Then
    ' USED to be sanity check, back when we only accepted docNumber/docYear format, like 5/18, but now also accepting new CARS format, words separated by space!
    'GSC_StandardName_ToFileName = ""
    'Exit Function
    suppliedStandardName = GSC_CarsNameFormat_ToStandardFormat(StandardName)    ' just transforming 1515 2018 into 1515/18, leaving suffix, doc type alone
Else
    suppliedStandardName = StandardName
End If

If suppliedStandardName = "" Then GSC_StandardName_ToFileName = "": Exit Function

' ELIMINATE superfluous suffix
If InStr(1, UCase(suppliedStandardName), "INIT") <> 0 Then
    suppliedStandardName = Replace(UCase(suppliedStandardName), " INIT", "")
End If

' If user provides language, use it
If Not IsMissing(DocLanguage) Then
    SLng = LCase(DocLanguage)
Else
    SLng = "xx"
End If

' Eliminate hard spaces and change multiple spaces to single, just in case
' (second case should not happend, but first probably will)
If InStr(1, Chr(160)) > 0 Then
    suppliedStandardName = UCase(Replace(suppliedStandardName, Chr(160), " "))  ' Chr(160) is "non-breaking space"
End If

' Now replace multiple spaces to single
Do While InStr(1, suppliedStandardName, "  ") > 0   ' Loop while two spaces
    ' InStr(1, "  ") > 0 Then
    suppliedStandardName = UCase(Replace(suppliedStandardName, "  ", " "))
Loop

If InStr(1, suppliedStandardName, " ") > 0 Then
    ' Eliminate spaces in suffixes if present
    suppliedStandardName = UCase(suppliedStandardName)
    
    Dim stNameSuffixesPart As String
    stNameSuffixesPart = Split(suppliedStandardName, "/")(1)
    
    stNameSuffixesPart = Replace(Replace(Replace(stNameSuffixesPart, "REV ", "REV"), "COR ", "COR"), "ADD ", "ADD")
    stNameSuffixesPart = Trim(Replace(Replace(Replace(stNameSuffixesPart, "AMD ", "AMD"), "EXT ", "EXT"), "DCL ", "DCL"))
                
    stNameSuffixesPart = Replace(Replace(Replace(stNameSuffixesPart, "RE ", "RE"), "CO ", "CO"), "AD ", "AD")
    stNameSuffixesPart = Trim(Replace(Replace(Replace(stNameSuffixesPart, "AM ", "AM"), "EX ", "EX"), "DC ", "DC"))
    
    ' and build back whole standard name, suffixes processed
    suppliedStandardName = Split(suppliedStandardName, "/")(0) & "/" & stNameSuffixesPart

    
    ' Replace short suffixes with long suffixes!
    ' for AD, we cannot contend with ANY place found for DOC TYPE ! It has to ne NOT at the beginning !
    If InStr(1, stNameSuffixesPart, "RE") > 0 And InStr(1, stNameSuffixesPart, "REV") = 0 Or _
        InStr(1, stNameSuffixesPart, "CO") > 0 And InStr(1, stNameSuffixesPart, "COR") = 0 Or _
        InStr(1, stNameSuffixesPart, "AD") > 1 And InStr(1, stNameSuffixesPart, "ADD") = 0 Or _
        InStr(1, stNameSuffixesPart, "AM") > 0 And InStr(1, stNameSuffixesPart, "AMD") = 0 Or _
        InStr(1, stNameSuffixesPart, "EX") > 0 And InStr(1, stNameSuffixesPart, "EXT") = 0 Or _
        InStr(1, stNameSuffixesPart, "DC") > 0 And InStr(1, stNameSuffixesPart, "DCL") = 0 Then
        
        ' Elongate suffixes back to original form !
        stNameSuffixesPart = Replace(Replace(Replace(stNameSuffixesPart, "RE", "REV"), "CO", "COR"), "AD", "ADD")
        stNameSuffixesPart = Trim(Replace(Replace(Replace(stNameSuffixesPart, "AM", "AMD"), "EX", "EXT"), "DC", "DCL"))
        
        ' In this situation (short WF suffixes), we have no space separation between diff suffixes present in compound suffixes block ! ("RE1CO1")
        stNameSuffixesPart = Replace(Replace(Replace(stNameSuffixesPart, "REV", " REV"), "COR", " COR"), "ADD", " ADD")
        stNameSuffixesPart = Trim(Replace(Replace(Replace(stNameSuffixesPart, "AMD", " AMD"), "EXT", " EXT"), "DCL", " DCL"))
        
        ' and build back whole standard name, suffixes processed
        suppliedStandardName = Split(suppliedStandardName, "/")(0) & "/" & stNameSuffixesPart
        
        ' Now replace multiple spaces to single
        Do While InStr(1, suppliedStandardName, "  ") > 0   ' Loop while two spaces
            ' InStr(1, "  ") > 0 Then
            suppliedStandardName = UCase(Replace(suppliedStandardName, "  ", " "))
        Loop
        
    End If

    
    StdNameArr = Split(Trim(suppliedStandardName), " ")   ' 0-based array, 'cause result of split function
Else
    ReDim StdNameArr(0)
    StdNameArr(0) = suppliedStandardName
End If



' Did the string contain spaces or not? (compulsory that it does, obviously)
If UBound(StdNameArr) = 0 Then
    
    If InStr(1, StdNameArr(0), "/") > 0 Then    ' have year, user supplied
        
        If UBound(Split(StdNameArr(0), "/")) = 2 Then   ' We have a revision number embedded within doc number, no suffix
            
            If IsNumeric(Split(StdNameArr(0), "/")(0)) And _
                IsNumeric(Split(StdNameArr(0), "/")(1)) And _
                IsNumeric(Split(StdNameArr(0), "/")(2)) And _
                Len(Split(StdNameArr(0), "/")(2)) = 2 Then
                
                    SType = "st"
                    SNumber = Format(Split(StdNameArr(0), "/")(0), "00000")
                    SYear = Format(Split(StdNameArr(0), "/")(2), "00")
                    SSuff = "re" & Format(Split(StdNameArr(0), "/")(1), "00")
                
            Else
                GSC_StandardName_ToFileName = ""    ' failure, these are necessary conditions
                Exit Function
            End If
        
        Else    ' No revision number embedded within doc number, no suffix
            ' Simplest of cases, one word supplied, doc number/ year, no suffixes
            If IsNumeric(Split(StdNameArr(0), "/")(0)) And _
                (IsNumeric(Split(StdNameArr(0), "/")(1)) And _
                Len(Split(StdNameArr(0), "/")(1)) = 2) Then
            
                    SType = "st"
                    SNumber = Format(Split(StdNameArr(0), "/")(0), "00000")
                    SYear = Format(Split(StdNameArr(0), "/")(1), "00")
                    SSuff = ""
                    
            Else
                GSC_StandardName_ToFileName = ""    ' failure, these are necessary conditions
                Exit Function
            End If
        End If
        
    Else    ' we add current year !
        
        SType = "st"
        SNumber = Format(StdNameArr(1), "00000")
        SYear = Right$(Trim(Year(Now)), 2)  ' two last digits of current year
        SSuff = ""  ' only two items is array, first being gsc doc type
        
    End If
    
Else    ' StdNameArr contains multiple items
    
    If CleanGSC_DocType(StdNameArr(0)) <> "" Then   ' First item is gsc doc type
        StdNameArr(0) = CleanGSC_DocType(StdNameArr(0))
        
        If UBound(StdNameArr) = 1 Then  ' Then there's no suffix
            
            ' Did the user supply year with number or not?
            If InStr(1, StdNameArr(1), "/") > 0 Then    ' have year, user supplied
                
                If UBound(Split(StdNameArr(1), "/")) = 2 Then   ' We have a revision number embedded within doc number
                    
                    If IsNumeric(Split(StdNameArr(1), "/")(0)) And _
                        IsNumeric(Split(StdNameArr(1), "/")(1)) And _
                        IsNumeric(Split(StdNameArr(1), "/")(2)) And _
                        (Len(Split(StdNameArr(1), "/")(2) = 2 Or Len(Split(StdNameArr(1), "/")(2)) = 4)) Then
                    
                            ERevisionNum = "re" & Format(Split(StdNameArr(1), "/")(1), "00")
                                                    
                            SType = LCase(StdNameArr(0))
                            SNumber = Format(Split(StdNameArr(1), "/")(0), "00000")
                            
                            If Len(Split(StdNameArr(1), "/")(2)) = 4 Then
                                SYear = Format(Right(Split(StdNameArr(1), "/")(2), 2), "00")
                            Else
                                SYear = Format(Split(StdNameArr(1), "/")(2), "00")
                            End If
                            
                            SSuff = ""  ' No suffix in array StdNameArr
                                                                            
                    Else
                    
                        GSC_StandardName_ToFileName = ""    ' failure, these are necessary conditions
                        Exit Function
                        
                    End If
                    
                Else    ' No embedded revision number
                
                    If IsNumeric(Split(StdNameArr(1), "/")(0)) And _
                        (IsNumeric(Split(StdNameArr(1), "/")(1)) And _
                        (Len(Split(StdNameArr(1), "/")(1)) = 2 Or Len(Split(StdNameArr(1), "/")(1)) = 4)) Then
                        
                            If Len(Split(StdNameArr(1), "/")(1)) = 4 Then
                                SYear = Format(Right(Split(StdNameArr(1), "/")(1), 2), "00")
                            Else
                                SYear = Format(Split(StdNameArr(1), "/")(1), "00")
                            End If
                            
                            SType = LCase(StdNameArr(0))
                            SNumber = Format(Split(StdNameArr(1), "/")(0), "00000")
                            
                            SSuff = ""
                            
                    Else
                        GSC_StandardName_ToFileName = ""    ' failure, these are necessary conditions
                        Exit Function
                    End If
                    
                End If
                
            Else    ' we add current year !
                
                SType = LCase(StdNameArr(0))
                SNumber = Format(StdNameArr(1), "00000")
                SYear = Right$(Trim(Year(Now)), 2)  ' two last digits of current year
                SSuff = ""  ' only two items is array, first being gsc doc type
                
            End If
        
        Else    ' There's at least one suffix
            
            If InStr(1, StdNameArr(1), "/") > 0 Then    ' have year, user supplied
                
                If UBound(Split(StdNameArr(1), "/")) = 2 Then   ' We have a revision number embedded within doc number
                        
                    If IsNumeric(Split(StdNameArr(1), "/")(0)) And _
                        IsNumeric(Split(StdNameArr(1), "/")(1)) And _
                        IsNumeric(Split(StdNameArr(1), "/")(2)) And _
                        Len(Split(StdNameArr(1), "/")(2)) = 2 Then
                    
                            ERevisionNum = "re" & Format(Split(StdNameArr(1), "/")(1), "00")
                                                    
                            SType = LCase(StdNameArr(0))
                            SNumber = Format(Split(StdNameArr(1), "/")(0), "00000")
                            SYear = Format(Split(StdNameArr(1), "/")(2), "00")
                                                                            
                    Else
                    
                        GSC_StandardName_ToFileName = ""    ' failure, these are necessary conditions
                        Exit Function
                        
                    End If

                        
                Else    ' No revision number embedded in doc number
                    If IsNumeric(Split(StdNameArr(1), "/")(0)) And _
                        (IsNumeric(Split(StdNameArr(1), "/")(1)) And _
                        (Len(Split(StdNameArr(1), "/")(1)) = 2 Or Len(Split(StdNameArr(1), "/")(1)) = 4)) Then
                            
                            ' allow for also using format from jobslip into this function (ie SN 1073/2015 ADD1)
                            If Len(Split(StdNameArr(1), "/")(1)) = 4 Then
                                SYear = Format(Right$(Split(StdNameArr(1), "/")(1), 2), "00")
                            Else
                                SYear = Format(Split(StdNameArr(1), "/")(1), "00")
                            End If
                            
                            SType = LCase(StdNameArr(0))
                            SNumber = Format(Split(StdNameArr(1), "/")(0), "00000")
                            
                            
                    Else
                        GSC_StandardName_ToFileName = ""    ' failure, these are necessary conditions
                        Exit Function
                    End If
                End If
                
            Else    ' we add current year !
                SType = LCase(StdNameArr(0))
                SNumber = StdNameArr(1)
                SYear = Right$(Trim(Year(Now)), 2)  ' two last digits of current year
            End If
            
            ' And now, gather suffixes
            For j = 1 To UBound(StdNameArr) - 1     ' elements 0 & 1 are excluded, we only count from there on
                If CleanGSC_DocSuffix(StdNameArr(j + 1)) <> "" Then
                    SSuff = SSuff & SGC_StdSuff_To_FileSuff(CleanGSC_DocSuffix(StdNameArr(j + 1)))
                Else
                    GSC_StandardName_ToFileName = ""    ' failure, these are necessary conditions
                    Exit Function
                End If
            Next j
            
            If ERevisionNum <> "" Then
                
            End If
            
        End If
        
    Else    ' gotta be number/year then, ha ?
        
        ' ******************************************************************************
        ' FIRST ITEM IS NUMBER/ YEAR, THERE'S AT LEAST ONE SUFFIX !
        ' ******************************************************************************
        
        ' Do we have user supplied year ? (Hopefully)
        If InStr(1, StdNameArr(0), "/") > 0 Then    ' have year, user supplied
            
            If UBound(Split(StdNameArr(0), "/")) = 2 Then   ' We have a revision number embedded within doc number, no suffix
                
                If IsNumeric(Split(StdNameArr(0), "/")(0)) And _
                    IsNumeric(Split(StdNameArr(0), "/")(1)) And _
                    IsNumeric(Split(StdNameArr(0), "/")(2)) And _
                    Len(Split(StdNameArr(0), "/")(2)) = 2 Then
                
                        ERevisionNum = "re" & Format(Split(StdNameArr(0), "/")(1), "00")
                                                
                        SType = "st"
                        SNumber = Format(Split(StdNameArr(0), "/")(0), "00000")
                        SYear = Format(Split(StdNameArr(0), "/")(2), "00")
                                                                        
                Else
                
                    GSC_StandardName_ToFileName = ""    ' failure, these are necessary conditions
                    Exit Function
                    
                End If
            
            Else    ' No revision number embedded in doc number
            
                If IsNumeric(Split(StdNameArr(0), "/")(0)) And _
                    (IsNumeric(Split(StdNameArr(0), "/")(1)) And _
                    Len(Split(StdNameArr(0), "/")(1)) = 2) Then
                
                        SType = "st"
                        SNumber = Format(Split(StdNameArr(0), "/")(0), "00000")
                        SYear = Format(Split(StdNameArr(0), "/")(1), "00")
                        
                Else
                    GSC_StandardName_ToFileName = ""    ' failure, these are necessary conditions
                    Exit Function
                End If
            
            End If
            
        Else    ' we add current year !
            SType = "st"
            SNumber = StdNameArr(0)
            SYear = Right$(Trim(Year(Now)), 2)  ' two last digits of current year
        End If
        
        ' And now, gather suffixes
        For j = 1 To UBound(StdNameArr)   ' elements 0 & 1 are excluded, we only count from there on
            
            If CleanGSC_DocSuffix(StdNameArr(j)) <> "" Then
                SSuff = SSuff & SGC_StdSuff_To_FileSuff(CleanGSC_DocSuffix(StdNameArr(j)))
            Else
                GSC_StandardName_ToFileName = ""    ' failure, these are necessary conditions
                Exit Function
            End If
        Next j
        
    End If

End If
    
' If first suffix is not a rev, we need add
If SSuff <> "" And Left(SSuff, 2) <> "re" Then
    If ERevisionNum <> "" Then
        SSuff = ERevisionNum & SSuff
    End If
ElseIf SSuff = "" And ERevisionNum <> "" Then
    SSuff = ERevisionNum
End If
    
' Did user request the file name result to have an extension or not ?
If Not IsMissing(WithFileExtension) Then
    If SSuff <> "" Then
        GSC_StandardName_ToFileName = SType & SNumber & "-" & SSuff & "." & SLng & SYear & ".doc"
    Else
        GSC_StandardName_ToFileName = SType & SNumber & "." & SLng & SYear & ".doc"
    End If
Else
    If SSuff <> "" Then
        GSC_StandardName_ToFileName = SType & SNumber & "-" & SSuff & "." & SLng & SYear
    Else
        GSC_StandardName_ToFileName = SType & SNumber & "." & SLng & SYear
    End If
End If


Exit Function


ErrorHandler:
    ' error handling code
End Function


Function GSC_StandardName_WithLang_toFileName(StandardName_andLanguage As String, Optional UseFileExtension, Optional IgnoreSuppliedLanguage) As String

' Input something like "SN 4196/1/13 EN" to obtain "sn04196-re01.en13.doc", with or without extension

' Also accepting Workflow Jobslip format for doc ID, such as "SN 2729/2018 REV1" or "ST 5964/2018 AD1RE4"

' CARS Search client and Automate format... WHY so many, oh, why?
' SN 1742 2018 REV6 RO    SN 1935 2018 REV1 RO   SN 2303 2018 ADD1 RO   SN 2381 2018 INIT RO   SN 2523 2018 INIT RO
' SN 2530 2018 INIT RO    SN 2559 2018 REV1 RO   SN 2560 2018 INIT RO   SN 2569 2018 INIT RO   SN 5099 2017 REV36 RO
' SN 10161 2018 INIT RO   ST 6291 2018 COR1 RO   ST 6583 2018 INIT RO   ST 6606 2018 COR1 RO   ST 6823 2018 INIT RO
' ST 6865 2018 INIT RO    ST 7084 2018 INIT RO   ST 7213 2018 INIT RO   ST 7231 2018 REV1 RO   ST 7310 2018 INIT RO
' ST 7323 2018 INIT RO    ST 7351 2018 INIT RO

' ALSO added optional parameter IgnoreSuppliedLanguage for successful parsing of email supplied projects to be checked cause deleted soon!


Dim pLang As String

Dim pGSCFileName As String

Dim suppliedDocID As String


If IsValidLangIso(Right$(StandardName_andLanguage, 2)) And _
    (Mid$(StandardName_andLanguage, Len(StandardName_andLanguage) - 2, 1) = " " Or Mid$(StandardName_andLanguage, Len(StandardName_andLanguage) - 2, 1) = ChrW(160)) Then ' also taking nbsp
    
    pLang = LCase(Right$(StandardName_andLanguage, 2)) ' Provided language
    
    ' Eliminate superfluous "no suffix" suffix !
    If InStr(1, UCase(StandardName_andLanguage), " INIT") > 0 Then
        suppliedDocID = Replace(UCase(StandardName_andLanguage), " INIT", "")     'INIT is tolerable but unnecessary, its a suffix meaning the doc does NOT have a suffix!
    Else
        suppliedDocID = StandardName_andLanguage
    End If
        
    pGSCFileName = GSC_StandardName_ToFileName(Trim(Left$(suppliedDocID, Len(suppliedDocID) - 2)))
    
    If pGSCFileName = "" Then
        GSC_StandardName_WithLang_toFileName = ""
        Exit Function
    Else
        If IsMissing(IgnoreSuppliedLanguage) Then
            GSC_StandardName_WithLang_toFileName = Replace(pGSCFileName, ".xx", "." & pLang)
        Else    ' user asked to still, ignore supplied language info. We still present xx!
            GSC_StandardName_WithLang_toFileName = pGSCFileName
        End If
    End If
    
Else
    ' We shall determine the original language from: the opened DW document's metadata or the original on M:Prod!
    GSC_StandardName_WithLang_toFileName = ""
End If


End Function


Sub StorageForCode()

' Two elements in array, first is gsc document type
If InStr(1, StdNameArr(1), "/") > 0 Then    ' have year, user supplied
    
    If IsNumeric(Split(StdNameArr(1), "/")(0)) And _
        (IsNumeric(Split(StdNameArr(1), "/")(1)) And _
        Len(Split(StdNameArr(1), "/")(1)) = 2) Then
    
            SType = UCase(StdNameArr(0))
            SNumber = Format(Split(StdNameArr(1), "/")(0), "00000")
            SYear = Format(Split(StdNameArr(1), "/")(1), "00")
            SSuff = ""
            
    Else
        GSC_StandardName_ToFileName = ""    ' failure, these are necessary conditions
        Exit Function
    End If
    
Else    ' we add current year !
    
    SType = StdNameArr(0)
    SNumber = StdNameArr(1)
    SYear = Right$(Trim(Year(Now)), 2)  ' two last digits of current year
    SSuff = ""  ' only two items is array, first being gsc doc type
    
End If

' Two elements in array, first is not gsc document type
If IsNumeric(Split(StdNameArr(0), "/")(0)) And _
    (IsNumeric(Split(StdNameArr(0), "/")(1)) And _
    Len(Split(StdNameArr(0), "/")(1)) = 2) Then

    If InStr(1, StdNameArr(0), "/") > 0 Then    ' have year, user supplied
        
        SType = "ST"
        SNumber = Format(Split(StdNameArr(0), "/")(0), "00000")
        SYear = Format(Split(StdNameArr(0), "/")(1), "00")
        
        ' We verify if its a GSC doc suffix, right ?
        If IsGSC_DocSuffix(StdNameArr(1)) Then
            SSuff = UCase(StdNameArr(1))
        Else
            ' What now ? Failure ?
            GSC_StandardName_ToFileName = ""
            Exit Function
        End If
        
    Else    ' we add current year !
        
        SType = "ST"
        SNumber = Format(StdNameArr(0), "00000")
        SYear = Right$(Trim(Year(Now)), 2)  ' two last digits of current year
        
        ' We verify if its a GSC doc suffix, right ?
        If IsGSC_DocSuffix(StdNameArr(1)) Then
            SSuff = UCase(StdNameArr(1))
        Else
            ' What now ? Failure ?
            GSC_StandardName_ToFileName = ""
            Exit Function
        End If
        
    End If
    
Else
    GSC_StandardName_ToFileName = ""    ' failure, these are necessary conditions
    Exit Function
End If


End Sub


Function CleanGSC_DocType(InputString As String) As String

' vers 0.2 - ADDED missing doc types

' Knows array of council document types and returns cleaned given input string if it's a council doc type, null if not


Dim gscDocTypesArr(21) As String

Dim icopy As String     ' input string copy

icopy = UCase(CleanString(Trim(Replace(Replace(InputString, vbCr, ""), Chr(160), ""))))


DocType(1, 0) = "1"
DocType(1, 1) = "2"

gscDocTypesArr(0) = "AC": gscDocTypesArr(1) = "AD": gscDocTypesArr(2) = "BU"
gscDocTypesArr(3) = "CG": gscDocTypesArr(4) = "CM": gscDocTypesArr(5) = "CP"
gscDocTypesArr(6) = "DA": gscDocTypesArr(7) = "DS": gscDocTypesArr(8) = "FA"
gscDocTypesArr(9) = "LT": gscDocTypesArr(10) = "NC": gscDocTypesArr(11) = "NP"
gscDocTypesArr(12) = "PE": gscDocTypesArr(13) = "RS": gscDocTypesArr(14) = "SN"
gscDocTypesArr(15) = "ST": gscDocTypesArr(16) = "XM": gscDocTypesArr(17) = "XN"
gscDocTypesArr(18) = "XT": gscDocTypesArr(19) = "EG": gscDocTypesArr(20) = "DE"
gscDocTypesArr(21) = "RE"

For i = 0 To UBound(gscDocTypesArr)
    If icopy = gscDocTypesArr(i) Then
        CleanGSC_DocType = icopy
        Exit Function
    End If
Next i

' if we reach this point without having jumped out, then function is false
CleanGSC_DocType = ""

End Function

Function CleanGSC_DocSuffix(InputString As String) As String
' Returns cleaned string if it is, empty string if not

Dim gscDocSuffArr(5) As String
Dim isCopy As String    ' Input string copy
Dim suffNum As String   ' gather suffix number in this, digit by digit

icopy = InputString

gscDocSuffArr(0) = "ADD": gscDocSuffArr(1) = "COR"
gscDocSuffArr(2) = "EXT": gscDocSuffArr(3) = "REV"
gscDocSuffArr(4) = "AMD": gscDocSuffArr(5) = "DCL"

' Need to clean string first, since removing rightside numbers might not work otherwise
icopy = UCase(CleanString(Trim(Replace(Replace(Replace(icopy, " ", ""), vbCr, ""), Chr(160), ""))))

' Gather into separate variable rightside numbers, if any,
' but remove them from icopy
Do While IsNumeric(Right$(icopy, 1))
    suffNum = suffNum & Right$(icopy, 1)
    icopy = Left$(icopy, Len(icopy) - 1)
Loop
suffNum = StrReverse(suffNum)

For i = 0 To UBound(gscDocSuffArr)
    If icopy = gscDocSuffArr(i) Then
        CleanGSC_DocSuffix = icopy & suffNum
        Exit Function
    End If
Next i

CleanGSC_DocSuffix = ""

End Function

Function SGC_Suff_ToStandard_Suff(suffix As String) As String
' version 0.4
' independent
' depends on Extr_Numbers
'
' Varianta cu spatiu a functiei de conversie a sufixului documentelor ("ADD 1")

Dim SufLTypes(5) As String
Dim SufSTypes(5) As String

SufLTypes(0) = "add"
SufLTypes(1) = "cor"
SufLTypes(2) = "ext"
SufLTypes(3) = "rev"
SufLTypes(4) = "amd"
SufLTypes(5) = "dcl"

SufSTypes(0) = "ad"
SufSTypes(1) = "co"
SufSTypes(2) = "ex"
SufSTypes(3) = "re"
SufSTypes(4) = "am"
SufSTypes(5) = "dc"

If Left$(suffix, 1) = "-" Then suffix = Mid$(suffix, 2, Len(suffix))

For i = 0 To 4
    If InStr(1, suffix, SufSTypes(i)) > 0 Then
        suffix = Replace(suffix, SufSTypes(i), " " & SufLTypes(i) & " ")
    End If
Next i

numsuf = Extr_Numbers(suffix)

For j = 0 To 3
    If numsuf(j) <> vbNothing Then
        suffix = Replace(suffix, numsuf(j), Format(numsuf(j), "0"))
    End If
Next j

SGC_Suff_ToStandard_Suff = Trim(UCase(suffix))

End Function

Function SGC_Suff_ToStandard_SuffNS(suffix As String) As String
' version 0.4
' independent
' depends on Extr_Numbers
'
' varianta fara spatiu a functiei SGC_Suff_ToStandard_SuffNS ("ADD1")

Dim SufLTypes(5) As String
Dim SufSTypes(5) As String

SufLTypes(0) = "add"
SufLTypes(1) = "cor"
SufLTypes(2) = "ext"
SufLTypes(3) = "rev"
SufLTypes(4) = "amd"
SufLTypes(5) = "dcl"

SufSTypes(0) = "ad"
SufSTypes(1) = "co"
SufSTypes(2) = "ex"
SufSTypes(3) = "re"
SufSTypes(4) = "am"
SufSTypes(5) = "dc"

If Left$(suffix, 1) = "-" Then suffix = Mid$(suffix, 2, Len(suffix))

For i = 0 To 4
    If InStr(1, suffix, SufSTypes(i)) > 0 Then
        suffix = Replace(suffix, SufSTypes(i), " " & SufLTypes(i))
    End If
Next i

numsuf = Extr_Numbers(suffix)

For j = 0 To 3
    If numsuf(j) <> vbNothing Then
        suffix = Replace(suffix, numsuf(j), Format(numsuf(j), "0"))
    End If
Next j

SGC_Suff_ToStandard_SuffNS = Trim(UCase(suffix))

End Function

Function SGC_StdSuff_To_FileSuff(suffix As String) As String    ' Long to short, spaced
' version 0.7
' independent
' depends on Extr_Numbers
'
' Varianta cu spatiu a functiei de conversie a sufixului documentelor ("ADD 1")

Dim SufLTypes(5) As String
Dim SufSTypes(5) As String
Dim alreadyDone As Boolean

Dim clsuff As String
clsuff = Replace(Replace(LCase(Trim(suffix)), " ", ""), Chr(160), "")

If Is_WFSuff(clsuff) Or Is_StdSuff(clsuff) Then
Else
    SGC_StdSuff_To_FileSuff = ""
    Exit Function
End If

SufLTypes(0) = "add"
SufLTypes(1) = "cor"
SufLTypes(2) = "ext"
SufLTypes(3) = "rev"
SufLTypes(4) = "amd"
SufLTypes(5) = "dcl"

SufSTypes(0) = "ad"
SufSTypes(1) = "co"
SufSTypes(2) = "ex"
SufSTypes(3) = "re"
SufSTypes(4) = "am"
SufSTypes(5) = "dc"

If Left$(suffix, 1) = "-" Then suffix = Mid$(suffix, 2, Len(suffix))

For i = 0 To UBound(SufSTypes)
    If InStr(1, LCase(suffix), SufLTypes(i)) > 0 Then
        suffix = Replace(LCase(suffix), SufLTypes(i), " " & SufSTypes(i) & " ")
    End If
Next i

numsuf = Extr_Numbers(suffix)

For j = 0 To 3
    If numsuf(j) <> vbNothing Then
        If j > 0 Then
            
            For k = j - 1 To 0 Step -1
                If numsuf(j) = numsuf(k) Then
                    alreadyDone = True
                End If
            Next k
            
        End If
        
        If Not alreadyDone Then
            suffix = Replace(Replace(suffix, numsuf(j), Format(numsuf(j), "00")), " ", "")
        End If
        alreadyDone = False
        
    End If
Next j

SGC_StdSuff_To_FileSuff = Trim(LCase(suffix))

End Function

Function Is_WFSuff(Phrase As String) As Boolean
' Variant of Is_StdSuff function, accepts two letters suffixes, such as RE instead of REV
' CO, AM, AD, RE & EX are all of the accepted ones.

Dim cph As String
cph = Replace(Replace(LCase(Trim(Phrase)), " ", ""), Chr(160), "")

If Not Is_StdSuff(Replace(Replace(Replace(Replace(Replace(Replace(Trim(LCase(cph)), "re", "rev"), "co", "cor"), "ex", "ext"), "am", "amd"), "ad", "add"), "dc", "dcl")) Then
    Is_WFSuff = False
Else
    Is_WFSuff = True
End If

End Function

Function Is_StdSuff(Phrase As String) As Boolean
'Function returns whether provided string is a combination of standard Council
'document suffixes, each followed by a digit or two. Standard Council suff are ADD, AMD,
'COR, EXT and REV. 0 alone is not allowed as it doesn't make sense to have REV0! (or the others)

Dim cleanedPhrase As String
cleanedPhrase = Trim(LCase(Replace(Replace(Phrase, " ", ""), Chr(160), "")))

'First check the length of provided cleaned (normalized) string
'since the criterias to come do not cover this situation
'(we could have a case such as REV234 identified as correct!)
If Len(cleanedPhrase) Mod 4 <> 0 And _
    Len(cleanedPhrase) Mod 5 <> 0 And _
    (Len(cleanedPhrase) < 4 * (Len(cleanedPhrase) \ 4) Or _
    Len(cleanedPhrase) > 5 * (Len(cleanedPhrase) \ 4)) Then

    Is_StdSuff = False
    Exit Function
End If

Dim numbersExtracted As String

numbersExtracted = Replace(Replace(Replace(cleanedPhrase, "add", "{#}"), "amd", "{#}"), "cor", "{#}")
numbersExtracted = Replace(Replace(Replace(numbersExtracted, "ext", "{#}"), "rev", "{#}"), "dcl", "{#}")
numbersExtracted = Replace(Replace(numbersExtracted, "{#}0", "{#}"), "{#}", "")
numbersExtracted = Trim(numbersExtracted)

Dim stdSuffNo As Integer
Dim expectedDigitsNo As Integer

stdSuffNo = (Len(cleanedPhrase) - Len(numbersExtracted)) \ 3
expectedDigitsNo = Len(cleanedPhrase) - 3 * stdSuffNo

If IsNumeric(numbersExtracted) And Len(numbersExtracted) = expectedDigitsNo Then
    Is_StdSuff = True
Else
    Is_StdSuff = False
End If

End Function



'Function OLDSGC_StdSuff_To_FileSuff(suffix As String) As String    ' Long to short, spaced
'' version 0.5
'' independent
'' depends on Extr_Numbers
''
'' Varianta cu spatiu a functiei de conversie a sufixului documentelor ("ADD 1")
'
'Dim SufLTypes(4) As String
'Dim SufSTypes(4) As String
'Dim alreadyDone As Boolean
'
'SufLTypes(0) = "add"
'SufLTypes(1) = "cor"
'SufLTypes(2) = "ext"
'SufLTypes(3) = "rev"
'SufLTypes(4) = "amd"
'
'SufSTypes(0) = "ad"
'SufSTypes(1) = "co"
'SufSTypes(2) = "ex"
'SufSTypes(3) = "re"
'SufSTypes(4) = "am"
'
'If Left$(suffix, 1) = "-" Then suffix = Mid$(suffix, 2, Len(suffix))
'
'For i = 0 To 4
'    If InStr(1, LCase(suffix), SufLTypes(i)) > 0 Then
'        suffix = Replace(LCase(suffix), SufLTypes(i), " " & SufSTypes(i) & " ")
'    End If
'Next i
'
'numsuf = Extr_Numbers(suffix)
'
'For j = 0 To 3
'    If numsuf(j) <> vbNothing Then
'        If j > 0 Then
'
'            For k = j - 1 To 0 Step -1
'                If numsuf(j) = numsuf(k) Then
'                    alreadyDone = True
'                End If
'            Next k
'
'        End If
'
'        If Not alreadyDone Then
'            suffix = Replace(Replace(suffix, numsuf(j), Format(numsuf(j), "00")), " ", "")
'        End If
'        alreadyDone = False
'
'    End If
'Next j
'
'SGC_StdSuff_To_FileSuff = Trim(LCase(suffix))
'
'End Function


Function Extr_LBNum(Phrase As String) As Double
' version 0.6
' independent
'
' This function extracts the longest number of digits from a string
' Found on net as "extract number sequence" and adapted.

Dim Length_of_String As Integer
Dim Current_Pos As Integer
Dim temp As String
Dim Found1 As Boolean
Dim BigstNumL As Long

Length_of_String = Len(Phrase)
temp = ""
Found1 = False
BigstNumL = 0

For Current_Pos = 1 To Length_of_String
    If (IsNumeric(Mid(Phrase, Current_Pos, 1))) = True Then
        temp = temp & Mid(Phrase, Current_Pos, 1)
        Found1 = True
    ElseIf (IsNumeric(Mid(Phrase, Current_Pos, 1))) = False And Found1 = True Then
'        If Len(Temp) = 1 Then
'            Found1 = False
'            Temp = ""
        If Len(temp) > BigstNumL Then
            BigstNumL = Len(temp)
            Found1 = False
            temp = ""
        End If
    End If
Next Current_Pos

'If Len(Temp) = 0 Then
'    Extr_LBNum = 0
'Else
'    If Len(Temp) > BigstNumL Then
'        BigstNumL = Len(Temp)
'        Extr_LBNum = BigstNumL
'    Else
'        Extr_LBNum = BigstNumL
'    End If
'End If

Extr_LBNum = BigstNumL

End Function


Function Extr_NumbersStruct(Phrase As String) As ExtractedNumbers

' version 0.4
' independent
'
' This function extracts the number sequences from a string

Dim extrNum As ExtractedNumbers


Dim Length_of_String As Integer
Dim Current_Pos, j As Integer
Dim temp As String
Dim Found1 As Boolean
Dim numbers()

Length_of_String = Len(Phrase)
temp = ""
Found1 = False
Let j = 0

For Current_Pos = 1 To Length_of_String
    
    If (IsNumeric(Mid(Phrase, Current_Pos, 1))) = True Then

        If Current_Pos = Length_of_String Then
            
            temp = temp & Mid(Phrase, Current_Pos, 1)
                        
            ReDim Preserve extrNum.numbers(j)
            extrNum.numbers(j) = temp
            extrNum.count = extrNum.count + 1
            
            Exit For
            
        Else


GoOn:       temp = temp & Mid(Phrase, Current_Pos, 1)
            Found1 = True
            
        End If
        
        
    ElseIf (IsNumeric(Mid(Phrase, Current_Pos, 1))) = False And Found1 = True Then
        If Mid(Phrase, Current_Pos, 1) = "." Then
            GoTo GoOn
        Else
LastOne:    Found1 = False
                        
            ReDim Preserve extrNum.numbers(j)
            
            extrNum.numbers(j) = temp
            extrNum.count = extrNum.count + 1
            
            j = j + 1
            
            temp = ""
        End If
    End If
Next Current_Pos


Extr_NumbersStruct = extrNum

End Function


Function Extr_Numbers(Phrase As String) As Variant
' version 0.6
' independent
'
' This function extracts the number sequences from a string
' Found on net and adapted it.

Dim Length_of_String As Integer
Dim Current_Pos, j As Integer
Dim temp As String
Dim Found1 As Boolean
Dim numbers(5)

Length_of_String = Len(Phrase)
temp = ""
Found1 = False
j = 0

For Current_Pos = 1 To Length_of_String
    If (IsNumeric(Mid(Phrase, Current_Pos, 1))) = True Then
GoOn:   temp = temp & Mid(Phrase, Current_Pos, 1)
        Found1 = True
    ElseIf (IsNumeric(Mid(Phrase, Current_Pos, 1))) = False And Found1 = True Then
        If Mid(Phrase, Current_Pos, 1) = "." Then
            GoTo GoOn
        Else
            Found1 = False
            j = j + 1
            numbers(j) = temp
            temp = ""
        End If
    End If
Next Current_Pos

j = j + 1
numbers(j) = temp

Extr_Numbers = Array(numbers(1), numbers(2), numbers(3), numbers(4))

End Function
Function Find_Subfolders()
' version 0.3
' secondary to Inject_SGCM_Folder
'

Dim mx2(), mx1, resmx
Dim fo As New Scripting.FileSystemObject
Dim bsf As folders
Dim UnuNuAre As Boolean
Dim fsDlgBaza As FileDialog               ' Initializam o variabila ca "FileDialog"
Dim fbazaS                                ' Calea directorului in care vom crea structura

' Instantiem variabila ca un filedialog de tipul "Folder Picker" - pentru a permite sa culegem de la utilizator calea
' unde doreste sa creeze structura de directoare
Set fsDlgBaza = Application.FileDialog(msoFileDialogFolderPicker)
fsDlgBaza.Title = "Rugam alegeti folderul radacina in care sa introducem metadatele"

' Verificam daca userul a selectat butonul de "actiune" (OK) si daca da, atribuim unei variabile rezultatul
' dialogului - in cazul nostru, folderul selectat.
If fsDlgBaza.Show = -1 Then
    fbazaS = fsDlgBaza.SelectedItems.item(1)
    'MsgBox fbazaS
Else
    ' Daca userul a selectat alt buton (Cancel), iesim direct
    End
End If

Set fbaza = fo.GetFolder(fbazaS)
If fbaza.SubFolders.count > 0 Then
    l = 0: Set bsf = fbaza.SubFolders
    For Each sf In bsf
        l = l + 1
        ReDim Preserve mx2(l)
        mx2(l) = sf.path
    Next sf
End If

Deepr:
For i = 0 To UBound(mx2)
    If fo.GetFolder(mx2(i)).SubFolders.count = 0 Then
        UnuNuAre = True
        Exit For
    End If
Next i

If UnuNuAre = False Then
    mx1 = mx2
    ReDim mx2(0)
Else
    resmx = mx2
    ' Pentru scopuri de debugging, in continuare construim lista de stringuri cu path-ul folderelor
    ' cautate.
    'For l = 1 To UBound(ResMx)
        'RsMs = RsMs & ResMx(l) & vbCr
    'Next l
    'Documents.Add.Content.InsertAfter ("Lista de foldere cautata a fost stabilita!" & vbCr & vbCr & RsMs)
    
    Find_Subfolders = resmx
    
    Exit Function
End If
    
k = 0
For j = 0 To UBound(mx1)
    Set fbaza = fo.GetFolder(mx1(j))
    For Each fs In fbaza.SubFolders
        k = k + 1
        ReDim Preserve mx2(k)
        mx2(k) = fs.path
    Next fs
Next j

GoTo Deepr

End Function
Function Sort_Array(iArray) As Variant

Dim strTemp As String
Dim x, y As Integer

For x = LBound(iArray) To (UBound(iArray) - 1)
    For y = (x + 1) To UBound(iArray)
        If iArray(x) > iArray(y) Then
            strTemp = iArray(x)
            iArray(x) = iArray(y)
            iArray(y) = strTemp
            strTemp = ""
        End If
    Next y
Next x
    
Sort_Array = iArray

End Function


' Bubble sort a bi-dimensional array ascending or descending, by first dimension or first and second dimension.
' Not possible to sort only by second dimension or sort ascending by one dimension and descending by another!
Function Sort_2D_Array(ByVal ArrayToSort As Variant, Optional SortDescending, Optional SortBySecondDimension) As Variant
         
    Dim tmpArray    ' to use during work
    tmpArray = ArrayToSort
    
    'Dim tmpArray(5, 2) As Variant
    'Dim v As Variant
    Dim i As Integer, j As Integer
    Dim r As Integer, c As Integer
    Dim temp As Variant
     
     'Create 2-dimensional array
     
'    v = Array(56, 22, "xyz", 22, 30, "zyz", 56, 30, "zxz", 22, 30, "zxz", 10, 18, "zzz", 22, 18, "zxx")
'    For i = 0 To UBound(v)
'        tmpArray(i \ 3, i Mod 3) = v(i)
'    Next
    
    Debug.Print "Unsorted array:"
    For r = LBound(tmpArray) To UBound(tmpArray)
        For c = LBound(tmpArray, 2) To UBound(tmpArray, 2)
            Debug.Print tmpArray(r, c);
        Next
        Debug.Print
    Next
    
    
     'Bubble sort column 0
    
    If IsMissing(SortDescending) Then
        ' Ascending
        For i = LBound(tmpArray) To UBound(tmpArray) - 1
            For j = i + 1 To UBound(tmpArray)
                If tmpArray(i, 0) > tmpArray(j, 0) Then
                    For c = LBound(tmpArray, 2) To UBound(tmpArray, 2)
                        temp = tmpArray(i, c)
                        tmpArray(i, c) = tmpArray(j, c)
                        tmpArray(j, c) = temp
                    Next
                End If
            Next
        Next
    Else     ' Descending
        For i = LBound(tmpArray) To UBound(tmpArray) - 1
            For j = i + 1 To UBound(tmpArray)
                If tmpArray(i, 0) < tmpArray(j, 0) Then
                    For c = LBound(tmpArray, 2) To UBound(tmpArray, 2)
                        temp = tmpArray(i, c)
                        tmpArray(i, c) = tmpArray(j, c)
                        tmpArray(j, c) = temp
                    Next
                End If
            Next
        Next
    End If
     
    
    If Not IsMissing(SortBySecondDimension) Then
        'Bubble sort column 1, where adjacent rows in column 0 are equal
        If IsMissing(SortDescending) Then
            ' Ascending
            For i = LBound(tmpArray) To UBound(tmpArray) - 1
                For j = i + 1 To UBound(tmpArray)
                    If tmpArray(i, 0) = tmpArray(j, 0) Then
                        If tmpArray(i, 1) > tmpArray(j, 1) Then
                            For c = LBound(tmpArray, 2) To UBound(tmpArray, 2)
                                temp = tmpArray(i, c)
                                tmpArray(i, c) = tmpArray(j, c)
                                tmpArray(j, c) = temp
                            Next
                        End If
                    End If
                Next
            Next
         Else    ' Descending
            For i = LBound(tmpArray) To UBound(tmpArray) - 1
                For j = i + 1 To UBound(tmpArray)
                    If tmpArray(i, 0) = tmpArray(j, 0) Then
                        If tmpArray(i, 1) < tmpArray(j, 1) Then
                            For c = LBound(tmpArray, 2) To UBound(tmpArray, 2)
                                temp = tmpArray(i, c)
                                tmpArray(i, c) = tmpArray(j, c)
                                tmpArray(j, c) = temp
                            Next
                        End If
                    End If
                Next
            Next
         End If
    End If
     
     'Output sorted array
'    Debug.Print "Sorted array:"
'    For r = LBound(tmpArray) To UBound(tmpArray)
'        For c = LBound(tmpArray, 2) To UBound(tmpArray, 2)
'            Debug.Print tmpArray(r, c);
'        Next
'        Debug.Print
'    Next
    
    Sort_2D_Array = tmpArray
    
End Function



Public Sub QuickSortArray(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1, Optional lngColumn As Long = 0)
    On Error Resume Next

    'Sort a 2-Dimensional array

    ' SampleUsage: sort arrData by the contents of column 3
    '
    '   QuickSortArray arrData, , , 3

    '
    'Posted by Jim Rech 10/20/98 Excel.Programming

    'Modifications, Nigel Heffernan:

    '       ' Escape failed comparison with empty variant
    '       ' Defensive coding: check inputs

    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim arrRowTemp As Variant
    Dim lngColTemp As Long

    If IsEmpty(SortArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then  'IsArray() is somewhat broken: Look for brackets in the type name
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortArray, 1)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray, 1)
    End If
    If lngMin >= lngMax Then    ' no sorting required
        Exit Sub
    End If

    i = lngMin
    j = lngMax

    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2, lngColumn)

    ' We  send 'Empty' and invalid data items to the end of the list:
    If IsObject(varMid) Then  ' note that we don't check isObject(SortArray(n)) - varMid *might* pick up a valid default member or property
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If

    While i <= j
        While SortArray(i, lngColumn) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j, lngColumn) And j > lngMin
            j = j - 1
        Wend

        If i <= j Then
            ' Swap the rows
            ReDim arrRowTemp(LBound(SortArray, 2) To UBound(SortArray, 2))
            For lngColTemp = LBound(SortArray, 2) To UBound(SortArray, 2)
                arrRowTemp(lngColTemp) = SortArray(i, lngColTemp)
                SortArray(i, lngColTemp) = SortArray(j, lngColTemp)
                SortArray(j, lngColTemp) = arrRowTemp(lngColTemp)
            Next lngColTemp
            Erase arrRowTemp

            i = i + 1
            j = j - 1
        End If
    Wend

    If (lngMin < j) Then Call QuickSortArray(SortArray, lngMin, j, lngColumn)
    If (i < lngMax) Then Call QuickSortArray(SortArray, i, lngMax, lngColumn)

End Sub



Function ReplaceForFind(stringtoconvert As String) As String

Dim SemneSuspecte(1, 8) As String

SemneSuspecte(0, 0) = "13"              'Carriage return character
SemneSuspecte(0, 1) = "10"              'LineFeed character
SemneSuspecte(0, 2) = "9"
SemneSuspecte(0, 3) = "11"              'Line Break Character
SemneSuspecte(0, 4) = "160"             'Hard (non-breaking) space
SemneSuspecte(0, 5) = "150"             'En Dash character
SemneSuspecte(0, 6) = "151"             'Em Dash character
SemneSuspecte(0, 7) = "30"              'Hard (non-breaking) hyphen
SemneSuspecte(0, 8) = "2"               'FootNote mark (daca textul titlului contine o referinta de nota
                                        'de subsol, dupa ce transferam in variabila acest text, variabila
                                        'va contine caracterul Chr(2), care, evident, trebuie inlaturat cu totul
SemneSuspecte(1, 0) = "^p"
SemneSuspecte(1, 1) = ""
SemneSuspecte(1, 2) = "^t"
SemneSuspecte(1, 3) = "^l"              'Nu am facut teste pt a sti daca Line Break ne incurca sau nu...
SemneSuspecte(1, 4) = "^s"
SemneSuspecte(1, 5) = "^="
SemneSuspecte(1, 6) = "^+"
SemneSuspecte(1, 7) = "^~"
SemneSuspecte(1, 8) = ""
        
FROnceMore:
If Right(stringtoconvert, 1) = vbCr Or Right(stringtoconvert, 1) = vbLf Then
    stringtoconvert = Left(stringtoconvert, (Len(stringtoconvert) - 1))
    GoTo FROnceMore
End If
                
For i = 0 To UBound(SemneSuspecte, 2)
    If InStr(1, stringtoconvert, Chr(SemneSuspecte(0, i))) > 0 Then
        stringtoconvert = Replace(stringtoconvert, Chr(SemneSuspecte(0, i)), SemneSuspecte(1, i))
    End If
Next i
      
ReplaceForFind = stringtoconvert
       
End Function
Sub RecunoasteDocConsiliu()
' version 0.99
' independent
' depends on VerificaCN
' 0.90 Modificat sa stabileasca variabilele SGCDocOrig si SGCDocRo chiar daca este rulata de pe un doc ro.
' 0.96 Modificat sa gaseasca limba doc-ului original in trei moduri (din bookmark-ul "LangueOrig", daca exista, din "recent files" din word, deschizand originalul cu GSC Menu "OpenDoc")
' 0.97
' 0.99 Reparat un bug ref la originalul care are ".COPY" in coada (deschis cu macroul "OpenDoc" al SGC)

Dim MsgS As String
Dim PointPos As Byte
Dim LinePos As Byte
Dim ff As New Scripting.FileSystemObject

If ActiveDocument.Variables.count < 2 Then
    GoTo startit
Else
    onehere = False: twohere = False
    For Each adv In ActiveDocument.Variables
        If adv.Name = "VSGCDocOr" Then
            onehere = True
        ElseIf adv.Name = "VSGCDocRo" Then
            twohere = True
        End If
    Next adv
        
    If onehere = True And twohere = True Then
        SGCDocOrig = ActiveDocument.Variables("VSGCDocOr").Value
        SGCDocRo = ActiveDocument.Variables("VSGCDocRo").Value
        If SGCDocOrig <> "" And SGCDocRo <> "" Then
            GoTo finally
        Else
            GoTo startit
        End If
    Else
        GoTo startit
    End If
End If


startit:
If Documents.count > 0 Then
    BaseDocName = ff.GetBaseName(ActiveDocument.Name)
    PointPos = InStr(1, BaseDocName, ".")
    Select Case PointPos
        Case 0
            MsgS = "Not Council document, or not properly named! Cherish the point!"
Ooops:      StatusBar = MsgS: Exit Sub
        Case Is = 8
            LinePos = InStr(1, BaseDocName, "-", vbTextCompare)
                If LinePos <> 0 Then MsgS = _
                "Not Council document, or not properly named! Where's your line at?": GoTo Ooops
        Case Is > 9
            If (PointPos - 9) Mod 4 <> 0 Then MsgS = _
            "Not Council document, or not properly named! Where's your point at?": GoTo Ooops
            LinePos = InStr(1, BaseDocName, "-", vbTextCompare)
            If LinePos <> 8 Then MsgS = _
            "Not Council document, or not properly named! Cherish your lines!": GoTo Ooops
        Case Else
            MsgS = "Not Council document, or not properly named! Where's your point at?": GoTo Ooops
    End Select
Else
    MsgS = "Would you please open a document?": GoTo Ooops
End If

Set ff = Nothing

DocType(1, 1) = "st": DocType(1, 8) = "cp"
DocType(1, 2) = "sn": DocType(1, 9) = "ad"
DocType(1, 3) = "cm": DocType(1, 10) = "ac"
DocType(1, 4) = "ds": DocType(1, 11) = "nc"
DocType(1, 5) = "lt": DocType(1, 12) = "np"
DocType(1, 6) = "bu": DocType(1, 13) = "da"
DocType(1, 7) = "pe": DocType(1, 14) = "rs"
DocType(1, 15) = "cg"

DocType(2, 1) = "1"
DocType(2, 2) = "2"

DocSuff(1, 1) = "re": DocSuff(1, 4) = "ex"
DocSuff(1, 2) = "co": DocSuff(1, 5) = "am"
DocSuff(1, 3) = "ad"

DocSuff(2, 1) = 8
DocSuff(2, 2) = PointPos - 8

ReDim DocYear(2, (Right$(DatePart("yyyy", Date), 2)))   'Ultimele doua cifre ale anului curent
j = 0

For i = Right$(DatePart("yyyy", Date), 2) To 1 Step -1
    j = j + 1
    DocYear(1, j) = Format(i, "00")
Next i

DocYear(2, 1) = PointPos + 3
DocYear(2, 2) = "2"

EUDocLang(1, 1) = "en": EUDocLang(1, 13) = "cs"
EUDocLang(1, 2) = "fr": EUDocLang(1, 14) = "et"
EUDocLang(1, 3) = "ro": EUDocLang(1, 15) = "lv"
EUDocLang(1, 4) = "de": EUDocLang(1, 16) = "lt"
EUDocLang(1, 5) = "nl": EUDocLang(1, 17) = "hu"
EUDocLang(1, 6) = "it": EUDocLang(1, 18) = "mt"
EUDocLang(1, 7) = "es": EUDocLang(1, 19) = "pl"
EUDocLang(1, 8) = "da": EUDocLang(1, 20) = "sk"
EUDocLang(1, 9) = "el": EUDocLang(1, 21) = "sl"
EUDocLang(1, 10) = "pt": EUDocLang(1, 22) = "bg"
EUDocLang(1, 11) = "fi": EUDocLang(1, 23) = "xx"
EUDocLang(1, 12) = "sv"

EUDocLang(2, 1) = PointPos + 1
EUDocLang(2, 2) = 2


For i = 1 To 5
    If VerificaCN(i, BaseDocName) = False Then
        MsgS = "Not Council Document, Sorry!"
        GoTo Ooops
    End If
Next i

Dim foundIt As Boolean
foundIt = False

For k = 1 To UBound(EUDocLang, 2)
    If InStr(1, BaseDocName, EUDocLang(1, k), vbTextCompare) <> 0 Then
        If EUDocLang(1, k) <> "ro" Then
            SGCDocOrig = BaseDocName & ".doc"
            SGCDocRo = Left$(BaseDocName, PointPos - 1) & Replace(BaseDocName, EUDocLang(1, k), "ro", PointPos, 1) & ".doc"
            Exit For
        Else
            ' tot restul codului, pina la iesirea din bucla for ("Next k") reprezinta de fapt incercarea de a gasi
            ' limba originala a documentului, fara sa incercam in retea decat ca ultima resursa (pentru laptopurile RUE)
            SGCDocRo = BaseDocName & ".doc"
            
            ' prima incercare este folosind bookmark-ul "LangueOrig", introdus de template-ul "_GenRo". Din pacate
            ' nu merge pentru CM-uri (si altele, cred)
            Dim iGotsIt As Boolean
            Dim trn As Range
            iGotsIt = False
            If ActiveDocument.Bookmarks.count > 0 Then
                For Each bm In ActiveDocument.Bookmarks
                    If bm.Name = "LangueOrig" Then
                        iGotsIt = True
                        Set trn = bm.Range
                        trn.MoveEndUntil (")")
                        SGCDocOrig = Replace(SGCDocRo, "ro", Mid$(trn.Text, 6, 2))
                        Exit For
                    End If
                Next bm
            End If
            
            ' incercam sa selectam documentul original printre documentele deschise, daca acesta este printre ele
            If iGotsIt = False Then
                If Application.Documents.count > 1 Then
                    For Each oDoc In Application.Documents
                        'Debug.Print oDoc.Name
                        If StrComp(oDoc.Name, SGCDocRo, vbTextCompare) <> 0 Then
                            For l = 1 To UBound(EUDocLang, 2)
                                If EUDocLang(1, l) <> "ro" Then
                                    whichLgOrig = Replace(SGCDocRo, "ro", EUDocLang(1, l))
                                    If StrComp(oDoc.Name, whichLgOrig) = 0 Or StrComp(oDoc.Name, (whichLgOrig & ".COPY")) = 0 Then
                                        SGCDocOrig = whichLgOrig
                                        iGotsIt = True
                                        Exit For
                                    End If
                                End If
                            Next l
                            If iGotsIt = True Then Exit For
                        End If
                    Next oDoc
                End If
            End If
                
            ' incercam deschiderea doc-ului original din "recent files" al wordului
            If iGotsIt = False Then
                If Application.RecentFiles.count > 0 Then
                    For Each rf In Application.RecentFiles
                        If rf.Name <> SGCDocRo Then
                            If Left$(rf.Name, 7) = Left$(SGCDocRo, 7) Then
                                For i = 1 To UBound(EUDocLang, 2)
                                    If Replace(SGCDocRo, "ro", EUDocLang(1, i)) = rf.Name Or _
                                    (Replace(SGCDocRo, "ro", EUDocLang(1, i)) & ".COPY") = rf.Name Then
                                        SGCDocOrig = rf.Name
                                        iGotsIt = True
                                        Exit For
                                    End If
                                Next i
                            End If
                            If iGotsIt = True Then Exit For
                        End If
                    Next rf
                End If
            End If
            
            ' incercam aducerea originalului folosind GSGMenu "OpenDoc"
            ' Documents.Add.Content.InsertAfter (SGCFileName_ToStandardName(SGCDocRo))
            If iGotsIt = False Then
                Documents.Add.Content.InsertAfter (Left$(SGCDocRo, 2) & " " & Mid$(SGCDocRo, 3, 5) & "/" & _
                Mid$(SGCDocRo, (InStr(1, SGCDocRo, ".") + 3), 2))
                'Debug.Print ActiveDocument.Name
                tmpdoc = ActiveDocument.Name
                OpenDoc.FindOriginalDoc
                tmpDoc1 = ActiveDocument.Name
                SGCDocOrig = Left$(ActiveDocument.Name, (Len(ActiveDocument.Name) - 5))
                Documents(SGCDocRo).Activate
            End If
            Exit For
        End If
    End If
Next k

' verificam existenta variabilelor care atesta recunoasterea doc-ului si, daca nu exista, le scrie dupa
' variabilele publice "SGCDocOrig" si "SGCDocRo"
On Error Resume Next
If ActiveDocument.Name = SGCDocRo Then
    ' daca documentul are variabile, facem "turul" acestora, pentru a le gasi pe ale noastre ("VSGCDocRo" si "VSGCDocOr"),
    ' daca exista
    If Documents(SGCDocRo).Variables.count >= 1 Then
        itshere = False
        alsohere = False
        For Each dv In Documents(SGCDocRo).Variables
            If dv.Name = "VSGCDocOr" Then
                itshere = True
            ElseIf dv.Name = "VSGCDocRo" Then
                alsohere = True
            End If
        Next dv
        
        If itshere = False Then
            Documents(SGCDocRo).Variables.Add "VSGCDocOr", SGCDocOrig
            Documents(SGCDocRo).Saved = False
        End If
        
        If alsohere = False Then
            Documents(SGCDocRo).Variables.Add "VSGCDocRo", SGCDocRo
            Documents(SGCDocRo).Saved = False
        End If
        If Documents(SGCDocRo).Saved = False Then Documents(SGCDocRo).Save
    ' daca nu are, le adaugam pur si simplu
    Else
        Documents(SGCDocRo).Variables.Add "VSGCDocOr", SGCDocOrig
        Documents(SGCDocRo).Variables.Add "VSGCDocRo", SGCDocRo
        Documents(SGCDocRo).Saved = False
        Documents(SGCDocRo).Save
    End If
End If

For Each d In Documents
    If d.Name = tmpdoc Then
        Documents(tmpdoc).Close SaveChanges:=wdDoNotSaveChanges
    ElseIf d.Name = tmpDoc1 Then
        Documents(tmpDoc1).Close SaveChanges:=wdDoNotSaveChanges
    End If
Next d

finally: StatusBar = "RecunoasteDocConsiliu: Council document verified as such. Variables established successfully."

End Sub
Sub PBrBef_ToPBr_Convert()
' Convert pagebreak before of paragraphs to manual pagebreaks at the same places

Dim p As Paragraph
Dim tr As Range

For Each p In ActiveDocument.Paragraphs
    'p.Range.Select
    If p.PageBreakBefore = True Then
        If p.Range.Information(wdWithInTable) = False Then
            p.PageBreakBefore = False
            
            Set tr = p.Range
            tr.Collapse wdCollapseStart
            tr.InsertBreak wdPageBreak
        End If
    End If
Next p

End Sub
Sub PBr_ToPRbBefore_Convert()
' Convert existing pagebreaks in document into "pagebreakbefore" setting
' applied to next paragraph following the page break character
' (thus producing exactly same result in document formatting)

Dim p As Paragraph
Dim cp As Paragraph
Dim tr As Range

For Each p In ActiveDocument.Paragraphs
    Set tr = p.Range
    'p.Range.Select
    If Asc(tr.Text) = 12 Then
        If p.Range.Characters.count > 1 Then
            p.PageBreakBefore = True
            tr.Collapse wdCollapseStart
            tr.Delete wdCharacter, 1
        End If
    End If
Next p

End Sub
Function sUserName() As String
sUserName = Environ$("username")
End Function
Function sComputerName() As String
sComputerName = Environ$("computername")
End Function

Function StdSufName(SufString As String) As String
' version 0.5
' independent
'

Dim SufLTypes(5) As String
Dim SufSTypes(5) As String

SufLTypes(1) = "add"
SufLTypes(2) = "cor"
SufLTypes(3) = "ext"
SufLTypes(4) = "rev"
SufLTypes(5) = "amd"

SufSTypes(1) = "ad"
SufSTypes(2) = "co"
SufSTypes(3) = "ex"
SufSTypes(4) = "re"
SufSTypes(5) = "am"

For i = 1 To 5
    If InStr(1, LCase(SufString), SufLTypes(i)) > 0 Then
        SufString = Replace(LCase(SufString), SufLTypes(i), SufSTypes(i))
    End If
Next i
StdSufName = LCase(SufString)

End Function


'/**
' * Returns True if input string contains at least one short (2 letters) form of suffix but not the long (3 ltrs) form of same
' */
Function Contains_Short_Suffixes(InputString) As Boolean

    If (InStr(1, InputString, "RE") > 0 And InStr(1, InputString, "REV") = 0) Or _
        (InStr(1, InputString, "CO") > 0 And InStr(1, InputString, "COR") = 0) Or _
        (InStr(1, InputString, "AD") > 0 And InStr(1, InputString, "ADD") = 0) Or _
        (InStr(1, InputString, "AM") > 0 And InStr(1, InputString, "AMD") = 0) Or _
        (InStr(1, InputString, "EX") > 0 And InStr(1, InputString, "EXT") = 0) Or _
        (InStr(1, InputString, "DC") > 0 And InStr(1, InputString, "DCL") = 0) Then
    
        Contains_Short_Suffixes = True
    
    End If


End Function

'/**
' * Function will return either a bidimensional array of strings with keys and values or a big string separated
' * with specified separators
' */
Public Function get_Listview_Values(Target_ListView_Control As ListView, Optional ReturnArray, Optional CategorySeparator, Optional ValuesSeparator) As Variant


If Target_ListView_Control.ListItems.count > 0 Then
    
    If Not IsMissing(ReturnArray) Then
        
        Dim resArr As Variant
        ReDim resArr(0, 0)
        
        For i = 1 To Target_ListView_Control.ListItems.count
            resArr(0, i - 1) = Target_ListView_Control.ListItems(i).Text
            resArr(1, i - 1) = Target_ListView_Control.ListItems(i).ListSubItems(1).Text
        Next i
        
        get_Listview_Values = resArr
        
    Else
        
        Dim catSep As String
        Dim valSep As String
        
        If Not IsMissing(CategorySeparator) Then
            catSep = CategorySeparator
        Else
            catSep = ":"
        End If
        
        If Not IsMissing(ValuesSeparator) Then
            valSep = ValuesSeparator
        Else
            valSep = ","
        End If
        
        Dim resString As String
        Dim lvLabels As String
        Dim lvValues As String
        
        For i = 1 To Target_ListView_Control.ListItems.count
            lvLabels = IIf(lvLabels = "", Target_ListView_Control.ListItems(i).Text, _
                lvLabels & valSep & Target_ListView_Control.ListItems(i).Text)
            lvValues = IIf(lvValues = "", Target_ListView_Control.ListItems(i).ListSubItems(1).Text, _
                lvValues & valSep & Target_ListView_Control.ListItems(i).ListSubItems(1).Text)
        Next i
        
        resString = lvLabels & catSep & lvValues
        
        get_Listview_Values = resString
        
    End If
    
Else
    get_Listview_Values = ""
End If

Exit Function

'**********************************************************************
errorGetLV:
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Function


Function Get_Index_ofArray_Entry(TargetArray, SoughtText As String) As Integer

For i = 0 To UBound(TargetArray)
    If StrComp(CStr(TargetArray(i)), SoughtText, vbTextCompare) = 0 Then
        Get_Index_ofArray_Entry = i
        Exit Function
    End If
Next i

Get_Index_ofArray_Entry = -1    ' error, sought text not found in supplied array

End Function

Function GetSize_ofFile(FileName As String, HostFolder As String) As Double


Dim sFile As Variant
Dim oShell: Set oShell = CreateObject("Shell.Application")
Dim oDir:   Set oDir = oShell.Namespace(CStr(HostFolder))


Set sFile = oDir.ParseName(FileName)

Dim sizeProp As String

sizeProp = oDir.GetDetailsOf(sFile, 1)


Select Case Split(sizeProp, " ")(1)
    
    Case "bytes"
        GetSize_ofFile = CDbl(Format(Split(sizeProp, " ")(0) / 1024, "0.00"))             ' bytes, we divide to get kilobytes
        
    Case "KB"
        GetSize_ofFile = CDbl(Format(Split(sizeProp, " ")(0), "0.00"))                    ' kilobytes, we read it directly
    
    Case "MB"
        GetSize_ofFile = CDbl(Format(Split(sizeProp, " ")(0) * 1024, "0.00"))             ' megabytes, we divide to get kilobytes
    
    Case "GB"
        GetSize_ofFile = CDbl(Format(Split(sizeProp, " ")(0) * 1024 * 1024, "0.00"))      ' bytes, we divide to get kilobytes
    
End Select


Set oShell = Nothing: Set oDir = Nothing: Set sFile = Nothing

End Function

Function TrimNonAlphaNums(InputString As String) As String

Dim istr As String
istr = InputString

Do While Len(istr) > 0 And (Not (Right$(istr, 1) Like "[a-zA-Z0-9)?]" Or Right$(istr, 1) Like "[]]"))
    istr = Left$(istr, Len(istr) - 1)
Loop

Do While Len(istr) > 0 And (Not (Left$(istr, 1) Like "[a-zA-Z0-9(]" Or Left$(istr, 1) Like "[[]"))
    istr = Mid$(istr, 2)
Loop

TrimNonAlphaNums = istr

End Function

Function GetArray_Index_forValue(ByRef TgArray, LookedUpValue As String) As Integer

'Debug.Print TgArray(0, 0)
'Debug.Print TgArray(1, 1)

For i = 0 To UBound(TgArray(0))
    If TgArray(0)(i) = LookedUpValue Then
        GetArray_Index_forValue = i
        Exit Function
    End If
Next i

GetArray_Index_forValue = -1    ' No found

End Function
