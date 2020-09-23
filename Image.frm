VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmBrowser 
   ClientHeight    =   7905
   ClientLeft      =   4725
   ClientTop       =   1410
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "Saturday Sans ICG"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   7905
   ScaleWidth      =   9675
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   7455
      Left            =   2160
      TabIndex        =   9
      Top             =   360
      Width           =   7410
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   -1  'True
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   5
      VideoBorderColor=   0
      VideoBorder3D   =   -1  'True
      Volume          =   0
      WindowlessVideo =   -1  'True
   End
   Begin VB.CheckBox chkFullScreen 
      Caption         =   "Full Screen Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Check and then click on the picture to see a full screen version."
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox txtFileName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2160
      MousePointer    =   3  'I-Beam
      TabIndex        =   7
      ToolTipText     =   "Type in new filename to rename the file"
      Top             =   0
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   240
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      ToolTipText     =   "Click to exit the program"
      Top             =   7560
      Width           =   615
   End
   Begin VB.DirListBox dirOne 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Double Click on a directory to display the files inside"
      Top             =   0
      Width           =   1815
   End
   Begin VB.DriveListBox drvOne 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Click to select a drive"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.FileListBox filOne 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Hidden          =   -1  'True
      Left            =   240
      MouseIcon       =   "Image.frx":0000
      MultiSelect     =   2  'Extended
      System          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Click on the thumbnail to see the full-sized image."
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Image imgOne 
      BorderStyle     =   1  'Fixed Single
      Height          =   7455
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   360
      Width           =   7455
   End
   Begin VB.Label lblFileSize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "file(s)."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label lblNumberOfFiles 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   7560
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileRefresh 
         Caption         =   "&Reload Directory"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewAll 
         Caption         =   "&All Images"
      End
      Begin VB.Menu mnuViewBitmap 
         Caption         =   "&Bitmap Images"
      End
      Begin VB.Menu mnuViewGif 
         Caption         =   "&GIF Images"
      End
      Begin VB.Menu mnuViewJpg 
         Caption         =   "&JPG Images"
      End
      Begin VB.Menu mnuMpeg 
         Caption         =   "&MPEG Clips"
      End
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iNumberOfFiles As Integer
Dim sFileName As String
Dim NumOfFiles As Integer
Dim CopyDestination As String
Dim currentfile As String
Dim FutureFile As String
Dim FileSize As Single
Dim CurrentFileSelected As Integer
Dim NewFileSelected As Integer
Dim Extension As String

Private Sub cmdExit_Click()
    End
End Sub

Private Sub dirOne_Change()
    'change the path for the file list box
    filOne.Path = dirOne.Path
        
    If filOne.ListCount = 0 Then
        frmBrowser.Width = 2250
    End If
    
    'calculate the number of picture files
    iNumberOfFiles = filOne.ListCount
    lblNumberOfFiles.Caption = iNumberOfFiles
End Sub

Private Sub drvOne_Change()
    'change the path in the directory list box
    dirOne.Path = drvOne
End Sub

Private Sub filOne_Click()
    'figure out the path and name of the file
    sFileName = filOne.Path & "\" & filOne.filename
    CurrentFileSelected = filOne.ListIndex
    frmBrowser.Width = 9795
    
    If Right(sFileName, 4) = "mpeg" Then
        txtFileName.Visible = True
        txtFileName.Enabled = False
        MediaPlayer1.Visible = True
        MediaPlayer1.Open sFileName
    ElseIf Right(sFileName, 4) <> "mpeg" Then
        chkFullScreen.Visible = True
        txtFileName.Enabled = True
        txtFileName.Visible = True
        MediaPlayer1.Visible = False
        imgOne.Visible = True
        'load the picture
        imgOne.Picture = LoadPicture(sFileName)
    End If
               
    NumOfFiles = filOne.ListIndex
    
    FileSize = FileLen(sFileName) / 1024
    lblFileSize.Caption = Format(FileSize, "standard") & " kb"
   
    txtFileName.Text = ""
    txtFileName.Text = filOne.filename
    currentfile = sFileName
End Sub

Private Sub filOne_DblClick()
    Call ShowImage
End Sub

Private Sub filOne_KeyDown(KeyCode As Integer, Shift As Integer)
    'if Enter is pressed display the picture
    If KeyCode = vbKeyReturn Then
        Call ShowImage
    ElseIf KeyCode = vbKeyDelete Then
        Kill sFileName
        filOne.Refresh
    End If
       
    NumOfFiles = filOne.ListIndex
End Sub

Private Sub Form_Load()
    Move 0, 0
    
    frmBrowser.Width = 2250
    imgOne.Visible = False
    MediaPlayer1.Visible = False
    chkFullScreen.Visible = False
    txtFileName.Visible = False
    
    dirOne.Path = App.Path
    filOne.Path = App.Path
    filOne.Refresh
    
    'make the file list box show only image files
    filOne.Pattern = "*.jpg;*.bmp;*.gif;*.jpeg;*.mpeg"
    
    'load the image form
    Load frmImage
    
    'calculate the files in the first directory
    iNumberOfFiles = filOne.ListCount
    lblNumberOfFiles.Caption = iNumberOfFiles
    
    Timer1.Enabled = True
    txtFileName.Enabled = False
End Sub

Private Sub imgOne_Click()
    Call ShowImage
End Sub

Private Sub mnuEditDelete_Click()
   Kill sFileName
   filOne.Refresh
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileRefresh_Click()
    filOne.Refresh
    iNumberOfFiles = filOne.ListCount
    lblNumberOfFiles.Caption = iNumberOfFiles
End Sub

Private Sub mnuMpeg_Click()
    filOne.Pattern = "*.mpeg;*.mpg"
    filOne.Refresh
    iNumberOfFiles = filOne.ListCount
    lblNumberOfFiles.Caption = iNumberOfFiles
End Sub

Private Sub mnuViewAll_Click()
    filOne.Pattern = "*.jpg;*.bmp;*.gif;*.jpeg;*.mpeg"
    filOne.Refresh
    iNumberOfFiles = filOne.ListCount
    lblNumberOfFiles.Caption = iNumberOfFiles
End Sub

Private Sub mnuViewBitmap_Click()
    filOne.Pattern = "*.bmp"
    filOne.Refresh
    iNumberOfFiles = filOne.ListCount
    lblNumberOfFiles.Caption = iNumberOfFiles
End Sub

Private Sub mnuViewGif_Click()
    filOne.Pattern = "*.gif"
    filOne.Refresh
    iNumberOfFiles = filOne.ListCount
    lblNumberOfFiles.Caption = iNumberOfFiles
End Sub

Private Sub mnuViewJpg_Click()
    filOne.Pattern = "*.jpg;*.jpeg"
    filOne.Refresh
    iNumberOfFiles = filOne.ListCount
    lblNumberOfFiles.Caption = iNumberOfFiles
End Sub

Private Sub Timer1_Timer()
    frmBrowser.Caption = dirOne.Path & "\" & filOne.filename
End Sub

Private Sub txtFileName_Click()
    txtFileName.SelStart = 0
    txtFileName.SelLength = 100
    NewFileSelected = CurrentFileSelected
    If Right(txtFileName.Text, 4) = "jpeg" Then
        Extension = Right(txtFileName.Text, 4)
    'ElseIf Right(txtFileName.Text, 4) = "mpeg" Then
    '    Extension = Right(txtFileName.Text, 4)
    Else
        Extension = Right(txtFileName.Text, 3)
    End If
End Sub

Private Sub txtFileName_LostFocus()
    FutureFile = filOne.Path & "\" & txtFileName.Text & "." & Extension
    Name currentfile As FutureFile
    filOne.Refresh
    filOne.Selected(NewFileSelected) = True
    
End Sub

Private Sub ShowImage()
    'load the picture
    frmImage!imgTwo.Picture = LoadPicture(sFileName)
    
    If chkFullScreen.Value = Checked Then
        frmImage!imgTwo.Stretch = True
        giHeight = Screen.Height
        giWidth = Screen.Width
    Else
        frmImage!imgTwo.Stretch = False
        giHeight = frmImage!imgTwo.Height
        giWidth = frmImage!imgTwo.Width
    End If
    
    'make the form the same size as the picture
    frmImage.Height = giHeight
    frmImage.Width = giWidth
    frmImage!imgTwo.Height = giHeight
    frmImage!imgTwo.Width = giWidth
    
    'show the picture and hide the browser
    frmImage.Show
    frmBrowser.Hide
    frmImage.Caption = filOne.filename
       
    NumOfFiles = filOne.ListIndex
End Sub
