VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExtractVB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extract images from VB files"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   Icon            =   "ExtractVB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6450
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cmnDialog 
      Left            =   4785
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Extract"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   75
      TabIndex        =   2
      Top             =   2865
      Width           =   1005
   End
   Begin VB.CommandButton frmCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   345
      Left            =   5355
      TabIndex        =   4
      Top             =   2865
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   3270
      TabIndex        =   6
      Top             =   0
      Width           =   3075
      Begin VB.PictureBox picGuage 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawMode        =   7  'Invert
         FillColor       =   &H000000FF&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   2760
         TabIndex        =   16
         Top             =   2385
         Width           =   2820
      End
      Begin VB.PictureBox picHolder 
         Height          =   720
         Left            =   120
         ScaleHeight     =   660
         ScaleWidth      =   2760
         TabIndex        =   14
         Top             =   1350
         Width           =   2820
         Begin VB.PictureBox picThumbnail 
            BorderStyle     =   0  'None
            Height          =   420
            Left            =   0
            ScaleHeight     =   420
            ScaleWidth      =   420
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   420
            Begin VB.Image imgThumbnail 
               Height          =   420
               Index           =   0
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Visible         =   0   'False
               Width           =   420
            End
         End
         Begin VB.HScrollBar HScroll 
            Enabled         =   0   'False
            Height          =   240
            Left            =   0
            TabIndex        =   3
            Top             =   420
            Width           =   2760
         End
      End
      Begin VB.PictureBox picImage 
         BackColor       =   &H80000005&
         Height          =   810
         Left            =   120
         ScaleHeight     =   50
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   184
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   450
         Width           =   2820
      End
      Begin VB.Label lblFile 
         Caption         =   "(Select files and destination folder)"
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   195
         Width           =   2820
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         Caption         =   "No images extracted"
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   2130
         UseMnemonic     =   0   'False
         Width           =   2820
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   75
      TabIndex        =   5
      Top             =   0
      Width           =   3075
      Begin VB.CommandButton cmdPickSource 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2685
         TabIndex        =   0
         Top             =   450
         Width           =   285
      End
      Begin VB.CommandButton cmdPickTarget 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2685
         TabIndex        =   1
         Top             =   2385
         Width           =   285
      End
      Begin VB.TextBox txtTarget 
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2385
         Width           =   2505
      End
      Begin VB.ListBox lstSource 
         Height          =   1620
         Left            =   105
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   450
         Width           =   2505
      End
      Begin VB.Label Label2 
         Caption         =   "Extract images from:"
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   195
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Destination folder:"
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   2130
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmExtractVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bProcessing As Boolean
Private bAbort As Boolean

Private sSource() As String
Private nImageCount As Integer

' ListBox Tooltips control
Private Declare Function SendLBMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_ITEMFROMPOINT As Long = &H1A9

Private Sub Form_Load()
    bProcessing = False
    bAbort = False
    nImageCount = 0
End Sub

Private Sub frmCancel_Click()
    If bProcessing Then
        bAbort = True
        DoEvents
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub lstSource_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 0 Then ' Only if no button was pressed
        Dim lXPoint As Long
        Dim lYPoint As Long
        Dim lIndex As Long
        '
        lXPoint = CLng(X / Screen.TwipsPerPixelX)
        lYPoint = CLng(Y / Screen.TwipsPerPixelY)
        '
        With lstSource
            ' Get selected item from list
            lIndex = SendLBMessage(.hWnd, LB_ITEMFROMPOINT, 0, ByVal ((lYPoint * 65536) + lXPoint))
            ' Show tip or clear last one
            If (lIndex >= 0) And (lIndex <= .ListCount) Then
                .ToolTipText = sSource(.ItemData(lIndex))
            Else
                .ToolTipText = ""
            End If
        End With
    End If
End Sub

Private Sub cmdPickSource_Click()
    On Error GoTo PickSourceCancelled

    With cmnDialog
        .DialogTitle = "Open VB files"
        .Filter = "VB files (*.frm;*.dob;*.ctl)|*.frm;*.dob;*.ctl|All files (*.*)|*.*"
        .FilterIndex = 1

        .CancelError = True
        .Flags = cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNFileMustExist Or cdlOFNExplorer
        .FileName = ""
        .MaxFileSize = 5120

        .ShowOpen

        .CancelError = False
    End With

    Dim sFileName As String
    sFileName = cmnDialog.FileName

    ' Build sSource() array
    Dim nCount As Integer, n As Integer
    Dim sFolder As String
    nCount = 0
    Erase sSource

    n = InStr(sFileName, Chr$(0))
    If n > 0 Then   ' Multi-select
        ' First one is the folder
        sFolder = Left$(sFileName, n - 1)
        sFileName = Mid$(sFileName, n + 1)
        ' The rest are the files
        Do While n > 0
            n = InStr(sFileName, Chr$(0))
            ReDim Preserve sSource(0 To nCount)
            If n = 0 Then
                sSource(nCount) = AttachPath(sFileName, sFolder)
            Else
                sSource(nCount) = AttachPath(Left$(sFileName, n - 1), sFolder)
                sFileName = Mid$(sFileName, n + 1)
            End If
            nCount = nCount + 1
        Loop
    Else            ' Single file...
        ReDim sSource(0)
        sSource(0) = sFileName
        nCount = 1
    End If

    ' Fill listbox
    Dim i As Integer
    With lstSource
        .Clear
        For i = 0 To (nCount - 1)
            .AddItem ExtractFileName(sSource(i))
            .ItemData(.NewIndex) = i
        Next
    End With

    Set_OK_State

PickSourceCancelled:
    cmnDialog.CancelError = False
End Sub

Private Sub cmdPickTarget_Click()
    Dim sFolder As String
    sFolder = FolderBrowser("Select destination folder for the images:", Me.hWnd)
    If sFolder <> "" Then
        txtTarget = sFolder
        Set_OK_State
    End If
End Sub

Private Sub HScroll_Change()
    picThumbnail.Left = -(HScroll.Value)
End Sub

Private Sub Set_OK_State()
    cmdOK.Enabled = (lstSource.ListCount > 0 And txtTarget <> "")
End Sub

Private Sub cmdOK_Click()
    ExtractImages
End Sub

Private Sub ExtractImages()
    Dim sFolder As String, sString As String, sUpper As String, _
            sImageData As String, sImageExt As String, sImageFile As String
    Dim sFileIn() As String
    Dim bScan As Boolean
    Dim nTotalSize As Long, nReadSize As Long   ' Total bytes to analyse (all files)
    Dim nProgress As Long
    Dim hIn As Integer, hOut As Integer, _
            i As Integer, nCount As Integer, _
            nInCount As Integer, _
            nSequence As Integer, n As Integer
    hIn = -1
    hOut = -1
    bAbort = False

    cmdOK.Enabled = False
    frmCancel.Caption = "Cancel"
    ProgressBar 0

    On Error GoTo ExtractError

    If nImageCount > 0 Then
        nImageCount = IIf(nImageCount > 78, 78, nImageCount)
        For i = nImageCount To 1 Step -1
            Unload imgThumbnail(i)
        Next
        picThumbnail.Width = 1
        HScroll.Enabled = False
        lblCount = "No images extracted"
    End If

    lblFile = "Checking source..."

    sFolder = txtTarget
    nCount = UBound(sSource)
    nInCount = 0
    nTotalSize = 0
    nReadSize = 0
    nImageCount = 0
    bScan = False

    ' Check of all files are available
    For i = 0 To nCount
        If FileExist(sSource(i)) Then
            ReDim Preserve sFileIn(0 To nInCount)
            sFileIn(nInCount) = sSource(i)
            nInCount = nInCount + 1
            nTotalSize = nTotalSize + FileLen(sSource(i))
        Else
            Select Case MsgBox("File '" & sSource(i) & "' does not exist! Do you wish to ignore this file during extraction, or check if file again?", vbAbortRetryIgnore + vbExclamation + vbDefaultButton3, "Missing File")
            Case vbAbort
                bAbort = True
                Exit For
            Case vbRetry
                i = i - 1
            End Select
        End If
    Next

    If bAbort Then GoTo ExtractExit

    If nInCount < 1 Then
        lblFile = "No files to analyse"
        MsgBox "There are no files to analyse. Please create a new list then try again.", vbExclamation, "No Files"
        Exit Sub
    End If

    lblFile = "Checking Target..."
    If Not FolderExist(sFolder) Then
        lblFile = "Invalid target folder"
        MsgBox "The target folder you specified is invalid. Please select another target folder.", vbExclamation, "Invalid Folder"
        Exit Sub
    End If

    lblFile = "Checks OK - Analysing"
    DoEvents: If bAbort Then GoTo ExtractExit   ' Yield to other processes - just in case Cancel is pressed

    For i = 0 To (nInCount - 1)

        DoEvents: If bAbort Then GoTo ExtractExit   ' Yield to other processes - just in case Cancel is pressed

        lblFile = "Analysing " & ExtractFileName(sFileIn(i))
        nSequence = 0
        sImageData = ""

        ' Open for for line-input...
        hIn = FreeFile
        Open sFileIn(i) For Input Access Read Shared As hIn

        Do While Not EOF(hIn)  ' Loop until end of file.

            DoEvents: If bAbort Then GoTo ExtractExit   ' Yield to other processes - just in case Cancel is pressed

            ' Update progressbar...
            nProgress = ((nReadSize + Seek(hIn)) * 100) / nTotalSize
            ProgressBar IIf(nProgress > 100, 100, nProgress)

            Line Input #hIn, sString
            sUpper = UCase$(Trim$(sString))

            If MatchString(sUpper, "ATTRIBUTE ") Then
                Exit Do
            ElseIf MatchString(sUpper, "BEGIN ") Then
                bScan = True
            ElseIf Not bScan Then
                GoTo EndOfFileLoop

            ElseIf MatchString(sUpper, "ICON ") Then
                '  Icon = "FormFile.frx":0000
                '       ^               ^
                n = InStr(sString, "=")
                If n > 0 Then
                    sString = Trim$(Mid$(sString, n + 1))
                    sImageData = ExtractImage(sString, sFileIn(i))
                    sImageExt = "ico"   ' Assume icon
                End If
            ElseIf MatchString(sUpper, "MOUSEICON ") Then
                '  MouseIcon = "FormFile.frx":0000
                '            ^               ^
                n = InStr(sString, "=")
                If n > 0 Then
                    sString = Trim$(Mid$(sString, n + 1))
                    sImageData = ExtractImage(sString, sFileIn(i))
                    sImageExt = "ico"   ' Assume icon (or cursor)
                End If
            ElseIf MatchString(sUpper, "PICTURE ") Then
                '  Picture = "FormFile.frx":0000
                '          ^               ^
                n = InStr(sString, "=")
                If n > 0 Then
                    sString = Trim$(Mid$(sString, n + 1))
                    sImageData = ExtractImage(sString, sFileIn(i))
                    sImageExt = "bmp"   ' Assume bitmap
                End If
            ElseIf MatchString(sUpper, "TOOLBOXBITMAP ") Then
                '  ToolboxBitmap = "FormFile.frx":0000
                '                ^               ^
                n = InStr(sString, "=")
                If n > 0 Then
                    sString = Trim$(Mid$(sString, n + 1))
                    sImageData = ExtractImage(sString, sFileIn(i))
                    sImageExt = "bmp"   ' Assume bitmap
                End If
            End If

            If sImageData <> "" Then
                ' Save image

                If Left(sImageData, 3) = "GIF" Then
                    sImageExt = "gif"
                ElseIf Left(sImageData, 2) = "BM" Then
                    sImageExt = "bmp"
                ElseIf Left(sImageData, 2) = (Chr$(0) & Chr$(0)) Then
                    sImageExt = "ico"
                ElseIf Mid(sImageData, 7, 4) = "JFIF" Then
                    sImageExt = "jpg"
                ElseIf Left(sImageData, 4) = "Î-ãÜ" Then
                    sImageExt = "wmf"
                End If

                sImageFile = AttachPath(ExtractFileName(sFileIn(i), False) & "_" & Format(nSequence, "000") & "." & sImageExt, sFolder)
                If FileExist(sImageFile) Then File2BAK sImageFile

                hOut = FreeFile
                Open sImageFile For Binary Access Write Lock Write As hOut
                Put #hOut, 1, sImageData
                Close hOut
                hOut = -1

                nSequence = nSequence + 1
                nImageCount = nImageCount + 1
                picImage = LoadPicture(sImageFile)

                ' Show in thumbnails...
                If nImageCount < 79 Then
                    Load imgThumbnail(nImageCount)
                    picThumbnail.Width = 420 * nImageCount
                    With imgThumbnail(nImageCount)
                        .Left = 420 * (nImageCount - 1)
                        .Picture = picImage.Picture
                        .ToolTipText = sImageFile
                        .Visible = True
                    End With

                    If picThumbnail.Width > HScroll.Width Then
                        HScroll.Enabled = True
                        With HScroll
                            .Max = (picThumbnail.Width - .Width)
                            .LargeChange = IIf(.Max < .Width, .Max, .Width)
                            If .Value > .Max Then
                                .Value = .Max
                                HScroll_Change
                            End If
                        End With
                        picThumbnail.SetFocus
                    End If
                End If

                lblCount = nImageCount & " image" & IIf(nImageCount = 1, "", "s") & " extracted"
                sImageData = ""
            End If

EndOfFileLoop:
        Loop

        nReadSize = nReadSize + LOF(hIn)
        Close hIn
    Next

    ProgressBar 100
    lblFile = "Extraction completed"

ExtractExit:
    On Error Resume Next
    If hIn > -1 Then Close hIn
    If hOut > -1 Then Close hOut
    frmCancel.Caption = "Close"
    Set_OK_State
    Exit Sub

ExtractError:
    MsgBox "Error occurred during extraction. Process aborted." & vbCrLf & _
            "(" & Err.Number & " - " & Err.Description & ")", vbCritical, "Extract Error"
    ProgressBar 0
    GoTo ExtractExit
End Sub

' Icon = "FormFile.frx":0000
'      ^               ^     = Markers
'        |-----------------| = Parameter
'
' Returns the image data in a string
'
Private Function ExtractImage(ByVal sString As String, sSourceFile As String) As String
    Dim n As Integer, nHandle As Integer
    Dim nOffset As Long, nFileSize As Long, nSize As Long
    Dim sFile As String, sData As String, sBytes As String
    Dim bFileOpen As Boolean

    bFileOpen = False

    On Error GoTo EI_ErrorHandler

    n = InStr(sString, ":")
    If n < 1 Then Exit Function

    sFile = AttachPath(StripQuotes(Left(sString, n - 1)), ExtractPath(sSourceFile))
    sString = "&H" & Trim$(Mid$(sString, n + 1))
    nOffset = Val(sString) + 1

    If Not FileExist(sFile) Then Exit Function

    nHandle = FreeFile
    Open sFile For Binary Access Read Shared As #nHandle
    bFileOpen = True
    nFileSize = LOF(nHandle)

    If (nOffset + 12) > nFileSize Then GoTo EI_ErrorHandler

    ' Get the header...
    Seek #nHandle, nOffset
    sData = Mid$(Input(12, #nHandle), 9, 4)

    ' Byte 9 to 12 (long) contains data size
    sBytes = "&H" & Right("00" & Hex(Asc(Mid$(sData, 4, 1))), 2) & _
            Right("00" & Hex(Asc(Mid$(sData, 3, 1))), 2) & _
            Right("00" & Hex(Asc(Mid$(sData, 2, 1))), 2) & _
            Right("00" & Hex(Asc(Mid$(sData, 1, 1))), 2)
    nSize = Val(sBytes)

    If nSize < 0 Or (nOffset + 11 + nSize) > nFileSize Then
        ' Try 28 byte header
        If (nOffset + 27) > nFileSize Then GoTo EI_ErrorHandler

        ' Get the header...
        Seek #nHandle, nOffset
        sData = Mid$(Input(28, #nHandle), 25, 4)

        ' Byte 25 to 28 (long) contains data size
        sBytes = "&H" & Right("00" & Hex(Asc(Mid$(sData, 4, 1))), 2) & _
                Right("00" & Hex(Asc(Mid$(sData, 3, 1))), 2) & _
                Right("00" & Hex(Asc(Mid$(sData, 2, 1))), 2) & _
                Right("00" & Hex(Asc(Mid$(sData, 1, 1))), 2)
        nSize = Val(sBytes)

        If nSize < 0 Or (nOffset + 27 + nSize) > nFileSize Then GoTo EI_ErrorHandler
    End If

    ' Get the data (position: nOffset + 13 - Already in position)
    ExtractImage = Input(nSize, #nHandle)

    ' That's it, the icon data is obtained
    Close #nHandle
    bFileOpen = False
    Exit Function

EI_ErrorHandler:
    If bFileOpen Then Close #nHandle
End Function

Private Function MatchString(sExpression As String, sContaining As String) As Boolean
    MatchString = (Left(sExpression, Len(sContaining)) = sContaining)
End Function

Private Function StripQuotes(ByVal sString As String) As String
    If Asc(Left(sString, 1)) = 34 And Asc(Right(sString, 1)) = 34 Then
        StripQuotes = Mid$(sString, 2, Len(sString) - 2)
    Else
        StripQuotes = sString
    End If
End Function

Private Sub ProgressBar(ByVal nPercent As Integer)
    Guage picGuage, nPercent
End Sub

' [Borrowed code below...]

Private Sub picGuage_Paint()
    ' this event will only get fired if AutoRedraw = False
    If IsNumeric(picGuage.Tag) = True Then Call Guage(picGuage, CInt(picGuage.Tag))
End Sub

Private Sub Guage(pic As Control, ByVal iPercent As Integer)
    ' this routine will draw a 3D guage in the PictureBox control
    ' pic is the control
    ' iPercent% is the percentage to show in the guage
    ' this is useful if you want to only show the guage when something is
    ' happening but not show it at other times
    ' the percentage to show will be stored into the Tag property so that
    ' we can tell what it is currently set to if we need to repaint it at
    ' a random time

    Dim sPercent$
    Dim iLeft%
    Dim iTop%
    Dim iRight%
    Dim iBottom%
    Dim iLineWidth%

    ' these are used to create the 3D effect
    Const DGREYCOLOUR& = &H808080
    Const LGREYCOLOUR& = &HC0C0C0
    Const WHITECOLOUR& = &HFFFFFF

    Const COPYPEN = 13
    Const XORPEN = 7

    ' validate our percentage
    If iPercent% < 0 Then
        iPercent% = 0
    ElseIf iPercent% > 100 Then
        iPercent% = 100
    End If

    ' set the number of twips per pixel into a variable
    ' NOTE: the picture control and the form it is on are expected to have
    ' their scale mode set to Twips
    iLineWidth% = Screen.TwipsPerPixelX

    ' I leave the BorderStyle set to 1 at design time so that the control is
    ' easy to find, but at run time we want the border to be invisible,
    ' however, just switching the border off will actually trigger a refresh
    ' of the control which is no use if AutoRedraw is set to False because
    ' that will trigger this code to run which will trigger another refresh
    ' which will ...
    If pic.BorderStyle <> 0 Then pic.BorderStyle = 0

    ' save the percentage into the Tag property - we can use this to repaint
    ' the guage if AutoRedraw is set to False
    pic.Tag = iPercent%

    ' set the text we will draw into a variable
    sPercent$ = CStr(iPercent%) & "%"

    ' work out the co-ords for the percentage bar
    iLeft% = iLineWidth%
    iTop% = iLineWidth%
    iRight% = pic.ScaleWidth - iLineWidth%
    iBottom% = pic.ScaleHeight - iLineWidth%

    ' erase everything by redrawing the background
    pic.DrawMode = COPYPEN
    pic.Line (iLeft%, iTop%)-(iRight%, iBottom%), pic.BackColor, BF

    ' add the text - work out where to put it first - nicely centered
    ' the default in VB3 is for bold text, change the FontBold property in
    ' the Picture control if you want this to be non-bold
    pic.CurrentX = (pic.ScaleWidth - pic.TextWidth(sPercent$)) / 2
    pic.CurrentY = (pic.ScaleHeight - pic.TextHeight(sPercent$)) / 2
    pic.Print sPercent$

    ' do the two colour bar by setting the DrawMode XOr then draw the bar
    ' in the fillcolour, if this overlaps the text then that portion of the
    ' text will get inverted, then XOr it again in the background colour,
    ' if you use the same colour for the FillColor and ForeColor then the
    ' text will invert nicely, but you can get some funny effects if you
    ' use two different colours
    ' NOTE: treat 0% as a special case because it will show up as a 1
    ' pixel wide line which looks bad
    ' ALSO NOTE: I am using BF in the call to the Line method, which means to
    ' draw a filled box, although I only want to draw lines which are a
    ' single pixel thick, because with trial and error I have found that this
    ' gives me the lines where I expect them for the co-ords that I am passing
    If iPercent% > 0 Then
        pic.DrawMode = XORPEN
        ' XOr the pen
        pic.Line (iLeft%, iTop%)-((iRight% / 100) * iPercent, iBottom%), pic.FillColor, BF
        pic.Line (iLeft%, iTop%)-((iRight% / 100) * iPercent, iBottom%), pic.BackColor, BF
    End If

    ' add the 3D look - right, bottom, top, left
    pic.DrawMode = COPYPEN
    pic.Line (iRight%, iLineWidth%)-(iRight%, iBottom%), WHITECOLOUR&, BF
    pic.Line (iLineWidth%, iBottom%)-(iRight%, iBottom%), WHITECOLOUR&, BF
    pic.Line (0, 0)-(iRight%, 0), DGREYCOLOUR&, BF
    pic.Line (0, 0)-(0, iBottom%), DGREYCOLOUR&, BF

    ' this line adds an additional grey border around the inside of the control to
    ' accentuate the 3D border - personal preference thing
    pic.Line (iLeft%, iTop%)-(iRight% - iLineWidth%, iBottom% - iLineWidth%), LGREYCOLOUR, B
End Sub
