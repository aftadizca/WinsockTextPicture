VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Client2 
   Caption         =   "CLIENT"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox tahunTB 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "Server1.frx":0000
      Left            =   6720
      List            =   "Server1.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2280
      Width           =   4935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton sendBT 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Caption         =   "SEND"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9000
      TabIndex        =   10
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CommandButton connectBT 
      Appearance      =   0  'Flat
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6480
      TabIndex        =   9
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox JurursanTB 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6720
      TabIndex        =   8
      Top             =   1800
      Width           =   4935
   End
   Begin VB.TextBox namaTB 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6720
      TabIndex        =   7
      Top             =   1320
      Width           =   4935
   End
   Begin VB.TextBox nimTB 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6720
      TabIndex        =   2
      Top             =   840
      Width           =   4935
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton loadBT 
      Appearance      =   0  'Flat
      Caption         =   "Load Image"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3585
      ScaleWidth      =   4545
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "BIODATA MAHASISWA"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   240
      Width           =   6735
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      Caption         =   "Jurusan"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4920
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      Caption         =   "Tahun Masuk"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4920
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Nama Mahasiswa"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4920
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "NIM"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4920
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "Client2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bSendingFile As Boolean
Private lTotal As Long
Public NumSockets As Integer

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim bFileArriving As Boolean
Dim sFile As String
Dim sArriving As String

Dim FileName As String 'nama
Dim FileTitle As String

Private Sub connectBT_Click()
    If connectBT.Caption = "Connect" Then
    Winsock1.Connect Winsock1.LocalIP, "11111"
    connectBT.Caption = "Conecting...."
    connectBT.Enabled = False
Else
    Winsock1.Close
    connectBT.Caption = "Connect"
End If
End Sub

Private Sub Form_Load()
    Dim tahun As Integer
    For tahun = 2010 To 2019
    tahunTB.AddItem tahun
    Next
    Picture1.Picture = LoadPicture(App.Path & "/foto.jpg")
    picStrech
    
End Sub


Private Sub loadBT_Click()
    CommonDialog1.Filter = "IMAGE|*.jpg"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        Picture1.Picture = LoadPicture(CommonDialog1.FileName)
        FileName = CommonDialog1.FileName
        FileTitle = GetGUID & ".jpg"
        picStrech
    End If
End Sub

'Function send file
Public Sub SendData(ByVal sFile As String, ByVal sSaveAs As String, ByVal msg As String, ByVal tcpCtl As Winsock)
'On Error GoTo ErrHandler
    Dim sSend As String, sBuf As String
    Dim ifreefile As Integer
    Dim lRead As Long, lLen As Long, lThisRead As Long, lLastRead As Long
    
    ifreefile = FreeFile
    
    If sFile <> "" Then
        tcpCtl.SendData "MSSGG" & msg
        DoEvents
        ' Open file for binary access:
        Open sFile For Binary Access Read As #ifreefile
        lLen = LOF(ifreefile)
        
        ' Loop through the file, loading it up in chunks of 64k:
        Do While lRead < lLen
            lThisRead = 65536
            If lThisRead + lRead > lLen Then
                lThisRead = lLen - lRead
            End If
            If Not lThisRead = lLastRead Then
                sBuf = Space$(lThisRead)
            End If
            Get #ifreefile, , sBuf
            lRead = lRead + lThisRead
            sSend = sSend & sBuf
        Loop
        lTotal = lLen
        Close ifreefile
        bSendingFile = True
        '// Send the file notification
        tcpCtl.SendData "FILE" & sSaveAs
        DoEvents
        '// Send the file
        tcpCtl.SendData sSend
        DoEvents
        '// Finished
        tcpCtl.SendData "FILEEND"
        bSendingFile = False
        Exit Sub
    Else
        tcpCtl.SendData "MSSGG" & msg
        DoEvents
    End If
'ErrHandler:
    'MsgBox "Errorssss " & Err & " : " & Error
End Sub

Private Sub sendBT_Click()
    SendData FileName, FileTitle, Label1.Caption & " : " & nimTB.Text & vbCrLf & Label2.Caption & " : " & namaTB.Text & vbCrLf & Label4.Caption & " : " & JurursanTB.Text & vbCrLf & Label3.Caption & " : " & tahunTB.Text & vbCrLf, Winsock1
End Sub

Private Sub Winsock1_Connect()
    connectBT.Enabled = True
    connectBT.Caption = "Disconnect"
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    Winsock1.Close
    Winsock1.Accept requestID
End Sub

Sub picStrech()
    'Picture1.Picture = Image1(1).Picture
    Picture1.ScaleMode = 3
    Picture1.AutoRedraw = True
    Picture1.PaintPicture Picture1.Picture, _
        0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, _
        0, 0, _
        Picture1.Picture.Width / 26.46, _
        Picture1.Picture.Height / 26.46
    Picture1.Picture = Picture1.Image
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    connectBT.Caption = "Connect"
    Winsock1.Close
    connectBT.Enabled = True
End Sub
