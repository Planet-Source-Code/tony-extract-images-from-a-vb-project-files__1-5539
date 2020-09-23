Attribute VB_Name = "modFileTools"
Option Explicit
' File(s) related procedures
' --------------------------

' Move a file API
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

' Copy a file
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long

' Delete to recycling bin API
Private Type SHFILEOPSTRUCT
   hWnd As Long
   wFunc As Long
   pFrom As String
   pTo As String
   fFlags As Integer
   fAborted As Boolean
   hNameMaps As Long
   sProgress As String
End Type
'
Private Const FO_DELETE As Long = &H3
Private Const FOF_ALLOWUNDO As Long = &H40
Private Const FOF_NOCONFIRMATION As Long = &H10
Private Const FOF_SILENT As Long = &H4
'
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

' File properties Constants and API
Private Type SHELLEXECUTEINFO
   cbSize As Long
   fMask As Long
   hWnd As Long
   lpVerb As String
   lpFile As String
   lpParameters As String
   lpDirectory As String
   nShow As Long
   hInstApp As Long
   lpIDList As Long     ' Optional parameter
   lpClass As String    ' Optional parameter
   hkeyClass As Long    ' Optional parameter
   dwHotKey As Long     ' Optional parameter
   hIcon As Long        ' Optional parameter
   hProcess As Long     ' Optional parameter
End Type
'
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
'
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long

' Browse for Folder Dialog
Public Type BrowseInfo
  hOwner As Long
  pIDLRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type
'
' BROWSEINFO.ulFlags values:
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
'
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long


' Checks wether file exist (handles wildcards too)
Function FileExist(ByVal sFile As String) As Boolean

    On Error GoTo ExistErrorHandler

    If Len(Trim(sFile)) = 0 Then
        ' Nothing given
        FileExist = False
        Exit Function
    ElseIf Right(sFile, 1) = "\" Or Right(sFile, 1) = ":" Then
        ' Just a part of a path or drive... (not complete)
        FileExist = False
        Exit Function
    ElseIf Dir(sFile) = "" Then
        ' Not there...
        FileExist = False
        Exit Function
    End If

    ' After all that torture, it must exist...
    On Error Resume Next
    FileExist = True
    Exit Function
ExistErrorHandler:
    On Error Resume Next
    FileExist = False
End Function

' Checks if file exist, and asks the user if it can be overwritten.
' With an option to give a new file name.
Function FileOverwriteDialog(ByRef sFile As String, oDialog As Object, Optional sFilter, Optional sDefaultExt) As Boolean
   If Not FileExist(sFile) Then
      FileOverwriteDialog = True
      Exit Function
   End If

   Select Case MsgBox(sFile + vbCrLf + "This file already exist." + vbCrLf + vbCrLf + "Replace existing file?", vbYesNoCancel + vbExclamation + vbDefaultButton3, "Save As")
   Case vbCancel
      FileOverwriteDialog = False
      Exit Function

   Case vbNo
      ' Pick new file
      On Error GoTo FileOverwriteCancelled

      With oDialog
         .DialogTitle = "Save file as ..."

         If IsMissing(sFilter) Then
            .Filter = "All files (*.*)|*.*"
         Else
            .Filter = sFilter
         End If
         .FilterIndex = 1

         If IsMissing(sDefaultExt) Then
            .DefaultExt = ""
         Else
            .DefaultExt = sDefaultExt
         End If

         .CancelError = True
         .flags = cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist

         .filename = sFile
      End With

      oDialog.ShowSave
      oDialog.CancelError = False
      sFile = oDialog.filename
   End Select

   FileOverwriteDialog = True
   Exit Function

FileOverwriteCancelled:
   oDialog.CancelError = False
   FileOverwriteDialog = False
End Function

' Extracts file extention from a file-string
Function ExtractFileExt(sFileName As String) As String
   Dim i As Integer
   For i = Len(sFileName) To 1 Step -1
      If InStr(".", Mid$(sFileName, i, 1)) Then Exit For
   Next
   ExtractFileExt = Right$(sFileName, Len(sFileName) - i)
End Function

Function ExtractFileName(ByVal sFileIn As String, Optional bIncludeExt As Boolean = True) As String
   Dim i As Integer
   For i = Len(sFileIn) To 1 Step -1
      If InStr(":\", Mid$(sFileIn, i, 1)) Then Exit For
   Next
   sFileIn = Mid$(sFileIn, i + 1, Len(sFileIn) - i)
   If Not bIncludeExt Then
      For i = Len(sFileIn) To 1 Step -1
         If InStr(".", Mid$(sFileIn, i, 1)) Then Exit For
      Next
      If i > 0 Then sFileIn = Left$(sFileIn, i - 1)
   End If
   ExtractFileName = sFileIn
End Function

' Extracts the path section of a file-string
Function ExtractPath(sPathIn As String) As String
   Dim i As Integer
   For i = Len(sPathIn) To 1 Step -1
      If InStr(":\", Mid$(sPathIn, i, 1)) Then Exit For
   Next
   ExtractPath = Left$(sPathIn, i)
End Function

' Adds a backslash if required
Function FixPath(ByVal sPath As String) As String
   If Len(Trim(sPath)) = 0 Then
      FixPath = ""
   ElseIf Right$(sPath, 1) <> "\" Then
      FixPath = sPath & "\"
   Else
      FixPath = sPath
   End If
End Function

' Concates file to a path (checks if a backslash is required)
Function AttachPath(sFileName As String, sPath As String) As String
   If Len(Trim(ExtractPath(sFileName))) = 0 Then
      AttachPath = FixPath(sPath) & sFileName
   Else
      AttachPath = sFileName
   End If
End Function

' Adds the application path to a filename
Function AppPathFile(sFileName As String) As String
   Dim sFullName As String
   sFullName = App.Path
   If Right$(sFullName, 1) <> "\" Then sFullName = sFullName & "\"
   AppPathFile = sFullName & sFileName
End Function

' Just in case the file and path name does not fit in a label.
Function LongDirFix(Incomming As String, Max As Integer) As String
   Dim i As Integer, LblLen As Integer, StringLen As Integer
   Dim TempString As String

   TempString = Incomming
   LblLen = Max

   If Len(TempString) <= LblLen Then
      LongDirFix = TempString
      Exit Function
   End If

   LblLen = LblLen - 6

   For i = Len(TempString) - LblLen To Len(TempString)
      If Mid$(TempString, i, 1) = "\" Then Exit For
   Next

   LongDirFix = Left$(TempString, 3) + "..." + Right$(TempString, Len(TempString) - (i - 1))
End Function

'-----------------------------------------------------------
' FUNCTION: FolderExist
'
' Determines whether the specified directory name exists.
' This function is used (for example) to determine whether
' an installation floppy is in the drive by passing in
' something like 'A:\'.
'
' IN: [sDirName] - name of directory to check for
'
' Returns: True if the directory exists, False otherwise
'-----------------------------------------------------------
'
Function FolderExist(ByVal sDirName As String) As Boolean
    Const WILDCARD As String = "*.*"

    Dim sDummy As String

    On Error Resume Next

    sDirName = FixPath(sDirName)
    sDummy = Dir$(sDirName & WILDCARD, vbDirectory)
    FolderExist = IIf(sDummy = "", False, True)

    Err = 0
End Function

'-----------------------------------------------------------
' FUNCTION: MakePath
'
' Creates the specified directory path.
' Uses MakePathAux() to create directories.
'
' In: [strDirName] - name of the dir path to make
'
' Returns: True if successful, False if error and the user chose to cancel.
'-----------------------------------------------------------
'
'Function MakePath(ByVal strDir As String) As Boolean
'   Do
'      If MakePathAux(strDir) Then
'         MakePath = True
'         Exit Function
'      Else
'         Dim strMsg As String
'         Dim iRet As Integer
'
'         strMsg = "Could not create directory:" & vbCrLf & strDir
'         iRet = MsgBox(strMsg, vbRetryCancel Or vbExclamation Or vbDefaultButton2, "Attention")
'         If iRet = vbCancel Then
'            MakePath = False
'            Exit Function
'         End If
'      End If
'   Loop
'End Function

'-----------------------------------------------------------
' FUNCTION: MakePathAux
'
' Creates the specified directory path.
'
' No user interaction occurs if an error is encountered.
' If user interaction is desired, use the related MakePath() function.
'
' In: [strDirName] - name of the dir path to make
'
' Returns: True if successful, False if error.
'-----------------------------------------------------------
'
Function MakePathAux(ByVal strDirName As String) As Boolean
    Const gstrSEP_DIR$ = "\"                         'Directory separator character
    
    Dim strPath As String
    Dim intOffset As Integer
    Dim intAnchor As Integer
    Dim strOldPath As String

    On Error Resume Next

    '
    'Add trailing backslash
    '
    If Right$(strDirName, 1) <> gstrSEP_DIR Then
        strDirName = strDirName & gstrSEP_DIR
    End If

    strOldPath = CurDir$
    MakePathAux = False
    intAnchor = 0

    '
    'Loop and make each subdir of the path separately.
    '
    '
    intOffset = InStr(intAnchor + 1, strDirName, gstrSEP_DIR)
    intAnchor = intOffset 'Start with at least one backslash, i.e. "C:\FirstDir"
    Do
        intOffset = InStr(intAnchor + 1, strDirName, gstrSEP_DIR)
        intAnchor = intOffset

        If intAnchor > 0 Then
            strPath = Left$(strDirName, intOffset - 1)
            ' Determine if this directory already exists
            Err = 0
            ChDir strPath
            If Err Then
                ' We must create this directory
                Err = 0
                MkDir strPath
            End If
        End If
    Loop Until intAnchor = 0

    MakePathAux = True
Done:
    ChDir strOldPath

    Err = 0
End Function

' Displays the Browse For Folder dialog and
' returns the folder that was chosen.
Function FolderBrowser(Optional sTitle As String = "Please select a folder:", Optional OwnerhWnd As Long = 0) As String
    Dim bInf As BrowseInfo
    Dim nPathID As Long
    Dim sPath As String
    Dim nOffset As Integer

    ' Set the properties of the folder dialog
    bInf.hOwner = OwnerhWnd
    bInf.lpszTitle = sTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS

    'Show the Browse For Folder dialog
    nPathID = SHBrowseForFolder(bInf)
    sPath = Space$(512)

    If SHGetPathFromIDList(ByVal nPathID, ByVal sPath) Then
        ' Trim off the null chars ending the path of the returned folder
        nOffset = InStr(sPath, Chr$(0))
        FolderBrowser = Left$(sPath, nOffset - 1)
    Else
        FolderBrowser = ""
    End If
End Function

' Creates a backup (.bak) file of given file
Sub File2BAK(sFile As String, Optional bKeepSource As Boolean = False)
   Dim i As Integer
   Dim sBAKFile As String
   Dim dl As Long

   If Not FileExist(sFile) Then Exit Sub

   sBAKFile = sFile
   For i = Len(sBAKFile) To 1 Step -1
      If InStr(".", Mid$(sBAKFile, i, 1)) Then
         sBAKFile = Left$(sBAKFile, i - 1)
         Exit For
      End If
   Next
   sBAKFile = sBAKFile & ".bak"

   If FileExist(sBAKFile) Then Kill sBAKFile

   If bKeepSource Then
      dl = CopyFile(sFile, sBAKFile, False)
   Else
      dl = MoveFile(sFile, sBAKFile)
   End If
End Sub

' Filenames with spaces in it, must be enclosed with quotes.
Function AddFileQuotes(ByVal sFile As String) As String
   sFile = Trim$(sFile)
   If InStr(sFile, " ") > 0 Then
      AddFileQuotes = Chr$(34) & sFile & Chr$(34)
   Else
      AddFileQuotes = sFile
   End If
End Function

Sub FileMove(sFile As String, sSourcePath As String, sTargetPath As String)
   Dim sSourceFile As String, sTargetFile As String

   sSourceFile = FixPath(sSourcePath) & sFile
   sTargetFile = FixPath(sTargetPath) & sFile

   On Error Resume Next
   Kill sTargetFile
   Call MoveFile(sSourceFile, sTargetFile)
End Sub

Sub FileRename(sSourceFile As String, sTargetFile As String)
    On Error Resume Next
    Kill sTargetFile
    Call MoveFile(sSourceFile, sTargetFile)
End Sub

Sub FileCopy(sSourceFile As String, sTargetFile As String)
    Dim bFailIfExists As Long
    bFailIfExists = CLng(False)

    Call CopyFile(sSourceFile, sTargetFile, bFailIfExists)
End Sub

' The function's ParamArray argument allows you to call it in several ways:
'
' Delete a single file:          lResult = ShellDelete("DELETE.ME")
'
' Pass file names in an array:   sFileName(1) = "DELETE.ME"
'                                sFileName(2) = "LOVE_LTR.DOC"
'                                sFileName(3) = "COVERUP.TXT"
'                                lResult = ShellDelete(sFileName())
'
' Pass file names as parameters: lResult = ShellDelete("DELETE.ME", "LOVE_LTR.DOC", "COVERUP.TXT")
'
Function ShellDelete(ParamArray vntFileName() As Variant) As Boolean
    Dim i As Integer, j As Integer
    Dim sFileNames As String
    Dim SHFileOp As SHFILEOPSTRUCT

    For i = LBound(vntFileName) To UBound(vntFileName)
        If IsArray(vntFileName(i)) Then
            For j = LBound(vntFileName(i)) To UBound(vntFileName(i))
                sFileNames = sFileNames & vntFileName(i)(j) & vbNullChar
            Next
        Else
            sFileNames = sFileNames & vntFileName(i) & vbNullChar
        End If
    Next
    sFileNames = sFileNames & vbNullChar

    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = sFileNames
        .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION Or FOF_SILENT
    End With

    Call SHFileOperation(SHFileOp)

    ShellDelete = SHFileOp.fAborted
End Function

Function ShellDeleteConfirm(fParent As Object, ParamArray vntFileName() As Variant) As Boolean
    Dim i As Integer, j As Integer
    Dim sFileNames As String
    Dim SHFileOp As SHFILEOPSTRUCT

    For i = LBound(vntFileName) To UBound(vntFileName)
        If IsArray(vntFileName(i)) Then
            For j = LBound(vntFileName(i)) To UBound(vntFileName(i))
                sFileNames = sFileNames & vntFileName(i)(j) & vbNullChar
            Next
        Else
            sFileNames = sFileNames & vntFileName(i) & vbNullChar
        End If
    Next
    sFileNames = sFileNames & vbNullChar

    With SHFileOp
        .hWnd = fParent.hWnd
        .wFunc = FO_DELETE
        .pFrom = sFileNames
        .fFlags = FOF_ALLOWUNDO
    End With

    Call SHFileOperation(SHFileOp)

    ShellDeleteConfirm = SHFileOp.fAborted
End Function

' Open a file properties property page for specified file if return value <=32 an error occurred
'
Function ShowFileProperties(sFileName As String, OwnerhWnd As Long) As Long
   Dim SEI As SHELLEXECUTEINFO
   Dim r As Long
        
   ' Fill in the SHELLEXECUTEINFO structure
   With SEI
      .cbSize = Len(SEI)
      .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
      .hWnd = OwnerhWnd
      .lpVerb = "properties"
      .lpFile = sFileName
      .lpParameters = vbNullChar
      .lpDirectory = vbNullChar
      .nShow = 0
      .hInstApp = 0
      .lpIDList = 0
   End With
 
   ' Call the API
   r = ShellExecuteEX(SEI)
 
   ' Return the instance handle as a sign of success
   ShowFileProperties = SEI.hInstApp
End Function
