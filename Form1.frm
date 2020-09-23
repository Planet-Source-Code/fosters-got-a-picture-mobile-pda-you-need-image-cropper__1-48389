VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Image Cropper"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7155
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   795
      Left            =   60
      TabIndex        =   37
      Top             =   7980
      Width           =   2655
      Begin VB.CommandButton Command5 
         Caption         =   "Exit"
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Image Manipulation"
      Enabled         =   0   'False
      Height          =   4635
      Left            =   60
      TabIndex        =   29
      Top             =   3300
      Width           =   2655
      Begin VB.CheckBox chkKeepProportion 
         Alignment       =   1  'Right Justify
         Caption         =   "&Keep proportions"
         Height          =   255
         Left            =   200
         TabIndex        =   6
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2220
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Resize Image"
         Height          =   375
         Left            =   1020
         TabIndex        =   7
         Top             =   1440
         Width           =   1395
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1740
         TabIndex        =   5
         Top             =   720
         Width           =   675
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1740
         TabIndex        =   4
         Top             =   360
         Width           =   675
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Crop Selection"
         Height          =   375
         Left            =   180
         TabIndex        =   11
         Top             =   3360
         Width           =   2235
      End
      Begin VB.TextBox txtSelHeight 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1740
         TabIndex        =   9
         Top             =   2520
         Width           =   675
      End
      Begin VB.TextBox txtSelWidth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1740
         TabIndex        =   8
         Top             =   2160
         Width           =   675
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Screen Dimensions"
         Height          =   375
         Left            =   180
         TabIndex        =   10
         Top             =   2880
         Width           =   2235
      End
      Begin VB.TextBox txtQuality 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Text            =   "90"
         Top             =   4140
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "JPEG Quality"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   4200
         Width           =   1395
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   10
         X2              =   2630
         Y1              =   3970
         Y2              =   3970
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "New Height"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   780
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "New Width"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   420
         Width           =   795
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   2640
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   10
         X2              =   2640
         Y1              =   1990
         Y2              =   1990
      End
      Begin VB.Label Label12 
         Caption         =   "Selection Width"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   2220
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Selection Height"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   2630
         Y1              =   3960
         Y2              =   3960
      End
   End
   Begin VB.CommandButton cmdMakeSource 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1515
   End
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Source Image"
      Height          =   1995
      Left            =   60
      TabIndex        =   20
      Top             =   1260
      Width           =   2655
      Begin VB.CommandButton cmdReload 
         Caption         =   "Reload"
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Open"
         Height          =   315
         Left            =   1380
         TabIndex        =   2
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblOrigWidth 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Left            =   1560
         TabIndex        =   26
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label lblOrigHeight 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Left            =   1560
         TabIndex        =   25
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Height"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Width"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Image Information:"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   1875
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Caption         =   "[file name]"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   300
         Width           =   2235
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "My Phone/PDA"
      Height          =   1155
      Left            =   60
      TabIndex        =   16
      Top             =   60
      Width           =   2655
      Begin VB.TextBox txtPhoneResY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Text            =   "320"
         ToolTipText     =   "Height"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtPhoneResX 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   780
         TabIndex        =   0
         Text            =   "240"
         ToolTipText     =   "Width"
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Height"
         Height          =   195
         Left            =   1380
         TabIndex        =   19
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Width"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   660
         Width           =   435
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Screen Resolution:"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   675
      Left            =   1380
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picDest 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   5040
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   14
      Top             =   540
      Width           =   1995
   End
   Begin VB.PictureBox picSource 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   2880
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   13
      Top             =   540
      Width           =   1995
   End
   Begin VB.PictureBox picDestBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   660
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   35
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Source"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   36
      Top             =   240
      Width           =   675
   End
   Begin VB.Image Image2 
      Height          =   195
      Left            =   780
      Picture         =   "Form1.frx":57E2
      Top             =   7320
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   2100
      Picture         =   "Form1.frx":658C
      Top             =   7440
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Dim m_cDib As New cDIBSection
Dim m_cDibBuffer As New cDIBSection
Dim sFile As String
Private CropSelectionActive As Boolean
Private SelectionColor As Long
Dim bKP As Boolean
Sub DrawRectangle(picIn As PictureBox, x1 As Long, y1 As Long, x2 As Long, y2 As Long, lColor As Long)
Dim hRPen As Long
    hRPen = CreatePen(0, 1, lColor)
    DeleteObject SelectObject(picIn.hdc, hRPen)
    Rectangle picIn.hdc, x1, y1, x2, y2
    DeleteObject hRPen
    DoEvents
End Sub

Private Sub cmdMakeSource_Click()
    picSource.Width = picDest.Width
    picSource.Height = picDest.Height
    picBuffer.Width = picDest.Width
    picBuffer.Height = picDest.Height
    
    picDest.Left = picSource.Left + picSource.Width + 120
    
    picSource.AutoRedraw = True
    BitBlt picSource.hdc, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picDest.hdc, 0, 0, vbSrcCopy
    BitBlt picBuffer.hdc, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picDest.hdc, 0, 0, vbSrcCopy
    picSource.Refresh
        
    m_cDib.Create picDest.ScaleWidth, picDest.ScaleHeight
    m_cDib.LoadPictureBlt picDest.hdc, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, vbSrcCopy

    lblOrigWidth = m_cDib.Width
    lblOrigHeight = m_cDib.Height
    
    SetActionButtonLocations
    
    If Me.WindowState <> 2 Then
        Me.Width = picDest.Left + picDest.Width + 250
        If (cmdSave.Left + cmdSave.Width + 250) > Me.Width Then Me.Width = (cmdSave.Left + cmdSave.Width + 250)
    End If
End Sub

Private Sub cmdReload_Click()
    LoadFile sFile
    cmdSave.Enabled = False
    cmdMakeSource.Enabled = False
End Sub

Private Sub cmdSave_Click()
Dim sI As String

   If VBGetSaveFileName(sI, , , "JPEG Files (*.JPG)|*.JPG|All Files (*.*)|*.*", , , , "JPG", Me.hwnd) Then

      If SaveJPG(m_cDib, sI, plQuality()) Then
         ' OK!
      Else
         MsgBox "Failed to save the picture to the file: '" & sI & "'", vbExclamation
      End If
   End If
End Sub
Private Function plQuality() As Long
   On Error Resume Next
   plQuality = CLng(txtQuality.Text)
   If Not Err.Number = 0 Then
      txtQuality.Text = "90"
      plQuality = 90
   End If
End Function
Private Sub Command1_Click()

    If VBGetOpenFileName(sFile, , , , , , "JPEG Files (*.JPG)|*.JPG|All Files (*.*)|*.*", 1, , , "JPG", Me.hwnd) Then
        LoadFile sFile
        cmdReload.Enabled = True
    Else
        MsgBox "There was a problem attempting to open this file", vbCritical, App.Title
        Exit Sub
    End If

End Sub
Sub LoadFile(sIN As String)
Dim lW As Long
Dim lH As Long
    
    picSource.Picture = LoadPicture(sFile)
    picDest.Picture = LoadPicture()
    DoEvents
    picDest.Move picSource.Left + picSource.Width + 120, picDest.TOp, picSource.Width, picSource.Height
    
    If Me.WindowState <> 2 Then
        If (picSource.Height + picSource.TOp + 600) > Me.Height Then
            Me.Height = picSource.Height + picSource.TOp + 600
        End If
        
        Me.Width = picDest.Left + picDest.Width + 250
        If (cmdSave.Left + cmdSave.Width + 250) > Me.Width Then Me.Width = (cmdSave.Left + cmdSave.Width + 250)
    End If
    picBuffer.Picture = picSource.Picture
    m_cDib.CreateFromPicture picSource
    
    lblOrigWidth = m_cDib.Width
    lblOrigHeight = m_cDib.Height
    
    If chkKeepProportion.Value = vbChecked Then
        bKP = True
        chkKeepProportion.Value = vbUnchecked
    Else
        bKP = False
    End If
    DoEvents
    txtWidth = IIf(IsNumeric(txtPhoneResX), txtPhoneResX, lblOrigWidth)
    txtHeight = IIf(IsNumeric(txtPhoneResX), txtPhoneResY, lblOrigHeight)
    If bKP Then chkKeepProportion.Value = vbChecked
    
    txtSelWidth = IIf(IsNumeric(txtPhoneResX), txtPhoneResX, lblOrigWidth)
    txtSelHeight = IIf(IsNumeric(txtPhoneResX), txtPhoneResY, lblOrigHeight)
    
    If picSource.Picture <> 0 Then Frame3.Enabled = True
    lblFile = Mid(sIN, InStrRev(sIN, "\") + 1)
    SetActionButtonLocations
    
End Sub
Sub SetActionButtonLocations()
    cmdMakeSource.Left = picDest.Left
    cmdSave.Left = cmdMakeSource.Left + cmdMakeSource.Width + 60
End Sub
Private Sub Command2_Click()
Dim cDib As New cDIBSection

Dim lW As Long
Dim lH As Long
    
    lW = CLng(txtWidth)
    lH = CLng(txtHeight)
    
    If (lW <> CLng(lblOrigWidth.Caption)) Or (lH <> CLng(lblOrigHeight.Caption)) Then
        picDest.AutoRedraw = True
        'm_cDib.CreateFromPicture picSource
        m_cDib.LoadPictureBlt picSource.hdc
        Set cDib = m_cDib.Resample(lH, lW)
        Set m_cDib = cDib
        m_cDibBuffer.Create m_cDib.Width, m_cDib.Height
        
        picDest.Width = m_cDib.Width * Screen.TwipsPerPixelX
        picDest.Height = m_cDib.Height * Screen.TwipsPerPixelY
        m_cDib.PaintPicture picDest.hdc
        picDest.Refresh
        
        DoEvents
        
        If Me.WindowState <> 2 Then
            If (picSource.Height + picSource.TOp + 600) > Me.Height Then
                Me.Height = picSource.Height + picSource.TOp + 600
            End If
    
            Me.Width = picDest.Left + picDest.Width + 250
            If (cmdSave.Left + cmdSave.Width + 250) > Me.Width Then Me.Width = (cmdSave.Left + cmdSave.Width + 250)
        End If

        cmdSave.Enabled = True
        cmdMakeSource.Enabled = True
        
        SetActionButtonLocations
        
    End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    picDest.AutoRedraw = False
    picSource.AutoRedraw = False
    CropSelectionActive = True
    SelectionColor = vbWhite

    picDest.Width = CLng(txtSelWidth) * Screen.TwipsPerPixelX
    picDest.Height = CLng(txtSelHeight) * Screen.TwipsPerPixelY
    picDestBuffer.Width = picDest.Width
    picDestBuffer.Height = picDest.Height
    
    If Me.WindowState <> 2 Then
        If (picSource.Height + picSource.TOp + 600) > Me.Height Then
            Me.Height = picSource.Height + picSource.TOp + 600
        End If

        Me.Width = picDest.Left + picDest.Width + 250
        If (cmdSave.Left + cmdSave.Width + 250) > Me.Width Then Me.Width = (cmdSave.Left + cmdSave.Width + 250)
    End If
    
End Sub

Private Sub Command4_Click()
    txtSelWidth = txtPhoneResX
    txtSelHeight = txtPhoneResY
    
End Sub

Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    App.Title = "Image Cropper"
    
    If Len(GetSetting(App.Title, "Settings", "PhoneResX")) = 0 Then
        SaveSetting App.Title, "Settings", "PhoneResX", 240
    End If
    txtPhoneResX = GetSetting(App.Title, "Settings", "PhoneResX")
    
    If Len(GetSetting(App.Title, "Settings", "PhoneResY")) = 0 Then
        SaveSetting App.Title, "Settings", "PhoneResY", 320
    End If
    txtPhoneResY = GetSetting(App.Title, "Settings", "PhoneResY")
    
    
    cmdSave.Picture = Image1.Picture
    cmdMakeSource.Picture = Image2.Picture


End Sub



Private Sub Form_Unload(Cancel As Integer)
    If IsNumeric(txtPhoneResX) And IsNumeric(txtPhoneResY) Then
        SaveSetting App.Title, "Settings", "PhoneResX", txtPhoneResX
        SaveSetting App.Title, "Settings", "PhoneResY", txtPhoneResY
    End If
End Sub

Private Sub picSource_Click()
    If Not CropSelectionActive Then Exit Sub

    CropSelectionActive = False
    BitBlt picSource.hdc, 0, 0, picBuffer.ScaleWidth, picBuffer.ScaleHeight, picBuffer.hdc, 0, 0, vbSrcCopy
    
    picDest.AutoRedraw = True
    BitBlt picDest.hdc, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, picDestBuffer.hdc, 0, 0, vbSrcCopy
    picDest.Refresh
    
    m_cDib.Create picDest.ScaleWidth, picDest.ScaleHeight
    m_cDib.LoadPictureBlt picDest.hdc, 0, 0, picDest.ScaleWidth, picDest.ScaleHeight, vbSrcCopy
    picDest.Refresh
    
    Command3.Enabled = True
    cmdSave.Enabled = True
    cmdMakeSource.Enabled = True
    SetActionButtonLocations
End Sub

Private Sub picSource_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lX As Long, lY As Long
Dim SW As Long, SH As Long
Dim lLeft As Long, lTop As Long
    If Not CropSelectionActive Then Exit Sub
    
    lX = CLng(x)
    lY = CLng(y)
    
    SW = CLng(txtSelWidth)
    SH = CLng(txtSelHeight)
    
    lTop = lY - (SH \ 2)
    If lTop < 0 Then lTop = 0
    If (lTop + SH) > picSource.ScaleHeight Then lTop = picSource.ScaleHeight - SH + 1
    
    lLeft = lX - (SW \ 2)
    If lLeft < 0 Then lLeft = 0
    If (lLeft + SW) > picSource.ScaleWidth Then lLeft = picSource.ScaleWidth - SW + 1
    
    SelectionColor = IIf(SelectionColor = 0, vbWhite, vbBlack)
    
    BitBlt picSource.hdc, 0, 0, CLng(lblOrigWidth), CLng(lblOrigHeight), picBuffer.hdc, 0, 0, vbSrcCopy
    DrawRectangle picSource, lLeft - 1, lTop - 1, lLeft + SW + 2, lTop + SH + 2, SelectionColor
    BitBlt picDestBuffer.hdc, 0, 0, SW, SH, picSource.hdc, lLeft, lTop, vbSrcCopy
    BitBlt picDest.hdc, 0, 0, SW, SH, picDestBuffer.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub txtHeight_Change()
Dim l As Long
    If (chkKeepProportion.Value = Checked) Then
        If (txtHeight.Tag = "") Then
            If IsNumeric(txtHeight) Then
                l = CLng(txtHeight)
                txtWidth.Tag = "txtHeight"
                txtWidth.Text = l * CLng(lblOrigWidth.Caption) \ CLng(lblOrigHeight.Caption)
                txtWidth.Tag = ""
            End If
        End If
    End If

End Sub


Private Sub txtWidth_Change()
Dim l As Long
    If (chkKeepProportion.Value = Checked) Then
        If txtWidth.Tag = "" Then
            If IsNumeric(txtWidth) Then
                l = CLng(txtWidth)
                txtHeight.Tag = "txtWidth"
                txtHeight.Text = l * CLng(lblOrigHeight.Caption) \ CLng(lblOrigWidth.Caption)
                txtHeight.Tag = ""
            End If
        End If
    End If
End Sub

