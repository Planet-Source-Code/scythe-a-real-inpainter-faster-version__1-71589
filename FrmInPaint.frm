VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmInPaint 
   Caption         =   "Real InPainting in VB"
   ClientHeight    =   7110
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7560
   Icon            =   "FrmInPaint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   474
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PicScroll 
      BorderStyle     =   0  'Kein
      Height          =   5775
      Left            =   120
      ScaleHeight     =   385
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   11
      Top             =   1320
      Width           =   7335
      Begin VB.VScrollBar VscrPic 
         Height          =   7335
         LargeChange     =   10
         Left            =   8520
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Frame FrmHide 
         BorderStyle     =   0  'Kein
         Height          =   255
         Left            =   8520
         TabIndex        =   13
         Top             =   7440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.HScrollBar HscrPic 
         Height          =   255
         LargeChange     =   10
         Left            =   0
         TabIndex        =   12
         Top             =   7560
         Visible         =   0   'False
         Width           =   8415
      End
      Begin VB.PictureBox PicOrg 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'Kein
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Ausgefüllt
         ForeColor       =   &H0000FF00&
         Height          =   5700
         Left            =   0
         MousePointer    =   2  'Kreuz
         OLEDropMode     =   1  'Manuell
         ScaleHeight     =   380
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   487
         TabIndex        =   15
         Top             =   0
         Width           =   7305
      End
   End
   Begin VB.CommandButton CmdBox 
      Caption         =   "Box"
      Height          =   495
      Left            =   5295
      Picture         =   "FrmInPaint.frx":1272
      Style           =   1  'Grafisch
      TabIndex        =   10
      Top             =   735
      Width           =   600
   End
   Begin VB.CommandButton CmdLine 
      Caption         =   "Poly"
      Height          =   495
      Left            =   4575
      Picture         =   "FrmInPaint.frx":15D6
      Style           =   1  'Grafisch
      TabIndex        =   9
      Top             =   735
      Width           =   600
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   10
      Left            =   6120
      Max             =   500
      Min             =   10
      SmallChange     =   2
      TabIndex        =   6
      Top             =   960
      Value           =   10
      Width           =   975
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   8
      Left            =   6120
      Max             =   32
      Min             =   2
      SmallChange     =   2
      TabIndex        =   3
      Top             =   240
      Value           =   4
      Width           =   975
   End
   Begin VB.CheckBox ChkPrev 
      Caption         =   "Preview"
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   420
      Value           =   1  'Aktiviert
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CMDlg 
      Left            =   3600
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton CmdInpaint 
      Caption         =   "Remove"
      Height          =   255
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label LblInfo2 
      Caption         =   $"FrmInPaint.frx":198C
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Shape ShpMark 
      BorderWidth     =   3
      Height          =   525
      Left            =   4560
      Top             =   720
      Width           =   630
   End
   Begin VB.Label Label5 
      Caption         =   "Scannborder"
      Height          =   255
      Left            =   6120
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "10"
      Height          =   255
      Left            =   7200
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Blocksize"
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   45
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "4"
      Height          =   255
      Left            =   7200
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.Label LblInfo 
      Caption         =   $"FrmInPaint.frx":1A84
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu MunLoad 
         Caption         =   "Load"
      End
      Begin VB.Menu MnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu MnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FrmInPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Simple form to demonstrate the Inpaint routine
'I´ve written the gui in about 10 Minutes
'the real thing is the inpainting.bas
'Scythe 2009
    
    
'Version 1.0.18
'Removed Bug with Picture Scrollbars
'The Scrollbars Position was on the Forms border and not on the Pictures Border
'
'Improved speed in Inpaint.bas
'Recoded PatchTexture to get 46% faster result with Bungee Jumper sample
'Thanks to ThePiper for his idea (about 17% more speed)


'Version 1.0.17
'Removed Compiling BUG
'In Project Properties / Compile / Advanced
'Disable Remove Array Bounds Checks
'If not the programm will crash if you inpaint near the borders of the picture
'
'Removed error in DoInPaint
'm_width = UBound(PicAr1, 1) should be m_width = UBound(PicAr1, 1) + 1
'm_height = UBound(PicAr1, 2) should be m_height = UBound(PicAr1, 2) + 1
'Now it scanns the whole picture
'
'Added Scrollbars to the Picture
'Now you can resize the form
'
'Added Box as Drawmode

'Fixed an error in Polydraw
'Added Close Poly if you click near the start Point
'
'Added Copy and Paste for fast transfer Picture to or from other Apps


    Const ABS_AUTOHIDE = &H1
    Const ABS_ONTOP = &H2
    Const ABM_GETSTATE = &H4
    Const ABM_GETTASKBARPOS = &H5
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type APPBARDATA
    cbSize As Long
    hwnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long '  message specific
End Type

Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Position() As POINTAPI
Private Polyctr As Long
Private DrawType As Long

Private Sub CmdInpaint_Click()
Dim X As Long
    If CmdInpaint.Caption = "Remove" Then
        CmdInpaint.Caption = "STOP"
        Me.MousePointer = 11
        StopIt = False
        X = DoInPaint(PicOrg, PicOrg, 0, 255, 0, CBool(ChkPrev.Value), HScroll1.Value, HScroll2.Value)
        CmdInpaint.Caption = "Remove"
        Me.MousePointer = 0
        If X > 0 Then MsgBox "Done" & vbCrLf & "Needed " & X & " repeats"
    Else
        CmdInpaint.Caption = "Remove"
        Me.MousePointer = 0
        StopIt = True
    End If
End Sub
Private Function InIDE() As Boolean

    On Error GoTo DivideError
    Debug.Print 1 / 0
    InIDE = False
    Exit Function
DivideError:
    InIDE = True

End Function

Private Sub CmdLine_Click()

    ShpMark.Top = CmdLine.Top - 1
    ShpMark.Left = CmdLine.Left - 1
    DrawType = 0
    LblInfo.Visible = True
    LblInfo2.Visible = False
    Polyctr = -1
    PicOrg.Cls

End Sub
Private Sub CmdBox_Click()

    ShpMark.Top = CmdBox.Top - 1
    ShpMark.Left = CmdBox.Left - 1
    DrawType = 1
    LblInfo.Visible = False
    LblInfo2.Visible = True
    Polyctr = -1
    PicOrg.Cls

End Sub

Private Sub Form_Load()

    If InIDE Then MsgBox "Compile me to see the real speed", vbCritical
    Polyctr = -1

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
'Strg C (Copy to Clipboard)
'Strg V (Paste from Clipboard)

    If KeyAscii = 22 Then
        If Clipboard.GetFormat(vbCFBitmap) = True Then
            PicOrg.Picture = Clipboard.GetData
            ResizePic
        End If
    End If
    If KeyAscii = 3 Then
        Clipboard.Clear
        Clipboard.SetData PicOrg.Image
    End If

End Sub
Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState = vbMaximized Then
        ResizePic
        Exit Sub
    End If
Dim X As Long
Dim Y As Long
    Y = Me.ScaleHeight

    X = Me.ScaleWidth
    If Y < 400 Then Y = 400
    Y = Y * Screen.TwipsPerPixelY + Me.Height - Me.ScaleHeight * Screen.TwipsPerPixelY
    If X < 504 Then X = 504
    X = X * Screen.TwipsPerPixelX + Me.Width - Me.ScaleWidth * Screen.TwipsPerPixelX
    Me.Width = X
    Me.Height = Y
    ResizePic

End Sub
Private Sub Form_Terminate()
    
    MnuExit_Click
   
End Sub
Private Sub ResizePic()

Dim X As Long
Dim Y As Long
    Y = (PicOrg.Height + PicScroll.Top + Me.Height / Screen.TwipsPerPixelY - Me.ScaleHeight + 4) * Screen.TwipsPerPixelY

    X = (PicOrg.Width + 24) * Screen.TwipsPerPixelX
    If X < 589 * Screen.TwipsPerPixelX Then X = 589 * Screen.TwipsPerPixelX
    If X > Screen.Width Then
        X = Screen.Width
    End If
    If Y > Screen.Height - TaskBarHeight Then
        Y = Screen.Height - TaskBarHeight
    End If
'Add Scrollbars if the picture is to big
    PicScroll.Width = Me.ScaleWidth - 16
    PicScroll.Height = Me.ScaleHeight - 88
    If PicOrg.Width > PicScroll.Width Then HscrPic.Visible = True Else HscrPic.Visible = False
    If PicOrg.Height > PicScroll.Height Then VscrPic.Visible = True Else VscrPic.Visible = False
    FrmHide.Visible = IIf(HscrPic.Visible Or VscrPic.Visible, True, False)
    X = IIf(PicScroll.Width - 12 > PicOrg.Width + 12, PicOrg.Width, PicScroll.Width - 12)
    Y = IIf(PicScroll.Height - 12 > PicOrg.Height + 12, PicOrg.Height, PicScroll.Height - 12)
    HscrPic.Move 0, Y, X, 12
    VscrPic.Move X, 0, 12, Y
    HscrPic.max = PicOrg.Width - HscrPic.Width
    VscrPic.max = PicOrg.Height - VscrPic.Height
    HscrPic.LargeChange = PicOrg.Width
    VscrPic.LargeChange = PicOrg.Height
    If Not HscrPic.Visible Then HscrPic.max = HscrPic.max - 12
    If Not VscrPic.Visible Then VscrPic.max = VscrPic.max - 12
    FrmHide.Move HscrPic.Width, VscrPic.Height, 12, 12
    HscrPic.Value = 0
    VscrPic.Value = 0

End Sub
Private Sub MunLoad_Click()

    On Error GoTo ErrOut
    CMDlg.Filter = "Pictures|*.bmp;*.gif;*.jpg"
    CMDlg.ShowOpen
    If CMDlg.filename <> "" Then
        Set PicOrg = LoadPicture(CMDlg.filename)
        ResizePic
    End If
ErrOut:

End Sub

Private Sub HScroll1_Change()

    Label2 = HScroll1.Value
    'for a good result the Scannborder has to be min 2 times the blocksize
    If HScroll2.Value < HScroll1.Value * 2 Then HScroll2.Value = HScroll1.Value * 2

End Sub

Private Sub HScroll2_Change()
Label4 = HScroll2.Value
End Sub

Private Sub MnuExit_Click()

    StopIt = True
    Unload Me
    End

End Sub
Private Sub MnuSave_Click()

    On Error GoTo ErrOut
    CMDlg.Flags = &H2
    CMDlg.Filter = "Windows Bitmap|*.bmp"
    CMDlg.ShowSave
    If CMDlg.filename <> "" Then
        SavePicture PicOrg.Image, CMDlg.filename
    End If
ErrOut:

End Sub

Private Sub PicOrg_DblClick()

    If DrawType <> 0 Then Exit Sub
    PicOrg.Cls
    PicOrg.AutoRedraw = True
    Polygon PicOrg.hdc, Position(0), Polyctr + 1
    PicOrg.AutoRedraw = False
    PicOrg.Refresh
    Polyctr = -1

End Sub
Private Sub PicOrg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim i As Long

     If Button = 2 Then
        Polyctr = -1
        PicOrg.Cls
        Exit Sub
    End If
    If DrawType = 1 Then
        If Polyctr = -1 Then
            Polyctr = 1
            ReDim Position(Polyctr)
            Position(Polyctr).X = X
            Position(Polyctr).Y = Y
            Else
            PicOrg.Cls
            PicOrg.AutoRedraw = True
            PicOrg.Line (Position(1).X, Position(1).Y)-(X, Y), , BF
            PicOrg.AutoRedraw = False
            Polyctr = -1
        End If
    Else
        If Polyctr > 1 Then
            If X > Position(0).X - 2 And X < Position(0).X + 2 And Y > Position(0).Y - 2 And Y < Position(0).Y + 2 Then
                PicOrg_DblClick
                Exit Sub
            End If
        End If
        Polyctr = Polyctr + 1
        ReDim Preserve Position(Polyctr)
        Position(Polyctr).X = X
        Position(Polyctr).Y = Y
        If Polyctr > 0 Then
            For i = 1 To Polyctr
                PicOrg.Line (Position(i - 1).X, Position(i - 1).Y)-(Position(i).X, Position(i).Y)
            Next i
        End If
    End If

End Sub
Private Sub PicOrg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim i As Long

    PicOrg.Cls
    If DrawType = 0 Then
        If Polyctr > -1 Then
            For i = 1 To Polyctr
                PicOrg.Line (Position(i - 1).X, Position(i - 1).Y)-(Position(i).X, Position(i).Y)
            Next i
            PicOrg.DrawMode = 6
            PicOrg.Line (Position(Polyctr).X, Position(Polyctr).Y)-(X, Y)
            PicOrg.DrawMode = 13
        End If
        Else
        If Polyctr = 1 Then
            PicOrg.DrawMode = 6
            PicOrg.Line (Position(1).X, Position(1).Y)-(X, Y), , BF
            PicOrg.DrawMode = 13
        End If
    End If

End Sub
Private Sub VscrPic_Change()

    PicOrg.Top = -VscrPic.Value

End Sub
Private Sub HscrPic_Change()

    PicOrg.Left = -HscrPic.Value

End Sub
Private Function TaskBarHeight() As Long

Dim ABD As APPBARDATA
Dim ret As Long
'Get the taskbar's position
    SHAppBarMessage ABM_GETTASKBARPOS, ABD
'Get the taskbar's state
    ret = SHAppBarMessage(ABM_GETSTATE, ABD)
    If (ret And ABS_AUTOHIDE) Then
        TaskBarHeight = 0
        Else
        TaskBarHeight = ABD.rc.Top
    End If

End Function

