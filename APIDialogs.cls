VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
    Persistable = 0  'NotPersistable
    DataBindingBehavior = 0  'vbNone
    DataSourceBehavior  = 0  'vbNone
    MTSTransactionMode  = 0	'NotAnMTSObject
END
Attribute VB_Name = "APIDialogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Enum FileOpenConstants
    cdlOFNReadOnly = &H1&
    cdlOFNOverwritePrompt = &H2&
    cdlOFNHideReadOnly = &H4&
    cdlOFNNoChangeDir = &H8&
    cdlOFNSHOWHELP = &H10&
    cdlOFNENABLEHOOK = &H20&
    cdlOFNENABLETEMPLATE = &H40&
    cdlOFNENABLETEMPLATEHANDLE = &H80&
    cdlOFNNoValidate = &H100&
    cdlOFNAllowMultiselect = &H200&
    cdlOFNExtensionDifferent = &H400&
    cdlOFNPathMustExist = &H800&
    cdlOFNFileMustExist = &H1000&
    cdlOFNCreatePrompt = &H2000&
    cdlOFNShareAware = &H4000&
    cdlOFNNoReadOnlyReturn = &H8000&
    cdlOFNNoDereferenceLinks = &H100000
    cdlOFNLongNames = &H200000
    cdlOFNNoLongNames = &H40000
    cdlOFNExplorer = &H80000
    cdlOFNDontAddToRecent = &H2000000
    cdlOFNForceShowHidden = &H10000000
End Enum
'
#If False Then ' Intellisense fix.
    Dim cdlOFNReadOnly, cdlOFNOverwritePrompt, cdlOFNHideReadOnly, cdlOFNNoChangeDir, cdlOFNSHOWHELP, cdlOFNENABLEHOOK, cdlOFNENABLETEMPLATE
    Dim cdlOFNENABLETEMPLATEHANDLE, cdlOFNNoValidate, cdlOFNAllowMultiselect, cdlOFNExtensionDifferent, cdlOFNPathMustExist, cdlOFNFileMustExist
    Dim cdlOFNCreatePrompt, cdlOFNShareAware, cdlOFNNoReadOnlyReturn, cdlOFNNoDereferenceLinks, cdlOFNLongNames, cdlOFNNoLongNames, cdlOFNExplorer, cdlOFNDontAddToRecent, cdlOFNForceShowHidden
#End If
'
Private Type OPENFILENAME
   lStructSize          As Long                 ' Length of structure, in bytes
   hwndOwner            As Long                 ' Window that owns the dialog, or NULL
   hInstance            As Long                 ' Handle of mem object containing template (not used)
   lpstrFilter          As Long                 ' File types/descriptions, delimited with vbnullchar, ends with 2xvbnullchar
   lpstrCustomFilter    As Long                 ' Filters typed in by user
   nMaxCustFilter       As Long                 ' Length of CustomFilter, min 40x chars
   nFilterIndex         As Long                 ' Filter Index to use (1,2,etc) or 0 for custom
   lpstrFile            As Long                 ' Initial file/returned file(s), delimited with vbnullchar for multi files
   nMaxFile             As Long                 ' Size of Initial File long  , min 256
   lpstrFileTitle       As Long                 ' File.ext excluding path
   nMaxFileTitle        As Long                 ' Length of FileTitle
   lpstrInitialDir      As Long                 ' Initial file dir, null for current dir
   lpstrTitle           As Long                 ' Title bar of dialog
   Flags                As FileOpenConstants    ' See FileOpenConstants
   nFileOffset          As Integer              ' Offset to file name in full path, 0-based
   nFileExtension       As Integer              ' Offset to file ext in full path, 0-based (excl '.')
   lpstrDefExt          As Long                 ' Default ext appended, excl '.', max 3 chars
   lCustData            As Long                 ' Appl defined data for lpfnHook
   lpfnHook             As Long                 ' Pointer to hook procedure
   lpTemplateName       As Long                 ' Template Name (not used)
   pvReserved           As Long                 ' new Win2000 / WinXP members
   dwReserved           As Long                 ' new Win2000 / WinXP members
   FlagsEx              As Long                 ' new Win2000 / WinXP members
End Type
'
Private Const cdlOfnFILE_OPEN_FLAGS = cdlOFNExplorer _
                                   Or cdlOFNLongNames _
                                   Or cdlOFNNoDereferenceLinks _
                                   Or cdlOFNHideReadOnly
Private Const cdlOfnFILE_SAVE_FLAGS = cdlOFNExplorer _
                                   Or cdlOFNLongNames _
                                   Or cdlOFNOverwritePrompt _
                                   Or cdlOFNHideReadOnly
'
Private Declare Function GetOpenFileNameW Lib "comdlg32" (lpofn As OPENFILENAME) As Long
Private Declare Function GetSaveFileNameW Lib "comdlg32" (lpofn As OPENFILENAME) As Long
'
Private OF As OPENFILENAME
'
' ShowOpen, ShowSave properties.
Public DefaultExt As String         ' Value excludes period.
Public DialogTitle As String
Public FileName As String
Public FullPath As String
Public FileTitle As String
Public Filter As String             ' Eg: "DLL files (*.dll)|*.dll|OCX files (*.ocx)|*.ocx"
Public FilterIndex As Long
Public InitDir As String
Public MaxFileSize As Long
Public OpenFlags As Long            ' See FileOpenConstants
Public SaveFlags As Long            ' See FileOpenConstants
'

Public Function ShowOpen(ByVal hWnd As Long) As Boolean
    ' Returns False on Cancel or error.
    ' hWnd:  The owner window.
    InitOpenFile hWnd
    OF.Flags = OpenFlags
    ShowOpen = GetOpenFileNameW(OF) <> 0
    If ShowOpen Then ExtractOpenFile
End Function

Public Function ShowSave(ByVal hWnd As Long) As Boolean
    ' Returns False on Cancel or error.
    ' hWnd:  The owner window.
    InitOpenFile hWnd
    OF.Flags = SaveFlags
    ShowSave = GetSaveFileNameW(OF) <> 0
    If ShowSave Then ExtractOpenFile
End Function

Private Sub InitOpenFile(ByVal hWnd As Long)
    
    ' change to allow for complete path to be available instead of just the file name
    If Len(FileName) < 1024 Then
        FileName = FileName & String$(1024 - Len(FileName), 0)
    End If
    
    With OF
        .lStructSize = LenB(OF) ' Ensure the structure size is set, change to allow for complete path to be available instead of just the file name
        .hwndOwner = hWnd
        .lpstrFilter = StrPtr(Replace$(Filter, "|", vbNullChar) & vbNullChar)
        .nFilterIndex = FilterIndex
        
        '.lpstrFile = StrPtr(FileName & String$(256 - Len(FileName), 0))
        .lpstrFile = StrPtr(FileName) ' change to allow for complete path to be available instead of just the file name
        .nMaxFile = MaxFileSize
        
        .lpstrFileTitle = StrPtr(String$(256, 0))
        .nMaxFileTitle = 256
        .lpstrInitialDir = StrPtr(InitDir)
        .lpstrTitle = StrPtr(DialogTitle)
        .lpstrDefExt = StrPtr(DefaultExt)
    End With
End Sub

Private Sub ExtractOpenFile()
    Dim i As Long
    With OF
        FileName = PtrToStrW(.lpstrFile)
        FileTitle = ""
        If (OpenFlags And cdlOFNAllowMultiselect) = 0 Then
            i = InStr(FileName, vbNullChar)
            If i Then FileName = Left$(FileName, i - 1)
            
            ' change to allow for complete path to be available instead of just the file name
            FullPath = FileName
            FileName = Mid$(FullPath, InStrRev(FullPath, "\") + 1)
            
            If (OpenFlags And cdlOFNNoValidate) = 0 Then
                FileTitle = PtrToStrW(.lpstrFileTitle)
                i = InStr(FileTitle, vbNullChar)
                If i Then FileTitle = Left$(FileTitle, i - 1)
            End If
        End If
        OpenFlags = .Flags
    End With
End Sub

Private Sub Class_Initialize()
    With OF
        .lStructSize = LenB(OF)
        .hInstance = App.hInstance
    End With
    Filter = "All files (*.*)|*.*"
    FilterIndex = 1
    MaxFileSize = 256
    DialogTitle = App.Title
    OpenFlags = cdlOfnFILE_OPEN_FLAGS
    SaveFlags = cdlOfnFILE_SAVE_FLAGS
End Sub

Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Function PtrToStrW(ByVal lpsz As Long) As String
    Dim lLen As Long
    If lpsz Then
        lLen = lstrlenW(lpsz)
        PtrToStrW = Space$(lLen)
        CopyMemory ByVal StrPtr(PtrToStrW), ByVal lpsz, lLen * 2
    End If
End Function