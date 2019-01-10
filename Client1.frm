VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Server2 
   Caption         =   "SERVER"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Listen"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4800
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Client1.frx":0000
      Top             =   720
      Width           =   6015
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
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "Server2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bSendingFile As Boolean
Private lTotal As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim bFileArriving As Boolean
Dim sFile As String
Dim sArriving As String

Dim FileName As String
Dim FileTitle As String

Private Sub Command1_Click()
If Command1.Caption = "Listen" Then
    Winsock1.LocalPort = "11111"
    Winsock1.Listen
    Command1.Caption = "Disconnect"
Else
    Winsock1.Close
    Command1.Caption = "Listen"
End If
End Sub

Private Sub Form_Load()
    Text1.Text = "NIM : " & vbCrLf & "Nama Mahasiswa: " & vbCrLf & "Jurusan : " & vbCrLf & "Tahun Masuk : " & vbCrLf
    Client2.Show
    Picture1.Picture = LoadPicture(App.Path & "/foto.jpg")
    picStrech
End Sub

Private Sub Winsock1_Close()
    Command1.Caption = "Listen"
    Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
    Command1.Enabled = True
    Command1.Caption = "Disconnect"
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    Winsock1.Close
    Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim ifreefile
    
    DoEvents
    Winsock1.GetData strData
    
    If Right$(strData, 7) = "FILEEND" Then
        bFileArriving = False
        sArriving = sArriving & Left$(strData, Len(strData) - 7)
        
        ifreefile = FreeFile
        
        If Dir(App.Path & "\tempClient\" & sFile) <> "" Then
            MsgBox "File Already Exists"
        Else
            Debug.Print App.Path & "\tempClient\" & sFile
            Open App.Path & "\tempClient\" & sFile For Binary Access Write As #ifreefile
            Put #ifreefile, 1, sArriving
            Close #ifreefile
            Picture1.Picture = LoadPicture(App.Path & "\tempClient\" & sFile)
            picStrech
            'ShellExecute 0, vbNullString, sFile, vbNullString, vbNullString, vbNormalFocus
            'RcvImg.Refresh
            'ChatDisplay.Text = ChatDisplay.Text & sFile & " received from " & Winsock1.RemoteHostIP & vbCrLf
        End If
        sArriving = ""
    ElseIf Left$(strData, 4) = "FILE" Then
        bFileArriving = True
        sFile = Right$(strData, Len(strData) - 4)
    ElseIf Left$(strData, 5) = "MSSGG" Then
        If Right$(strData, Len(strData) - 5) <> "" Then
            Text1.Text = Right$(strData, Len(strData) - 5)
            'ChatDisplay.Text = ChatDisplay.Text & Winsock1.RemoteHostIP & " : " & Right$(strData, Len(strData) - 5) & vbCrLf
        End If
        bFileArriving = False
    ElseIf bFileArriving Then
        sArriving = sArriving & strData
    End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Command1.Caption = "Listen"
    Winsock1.Close
    Command1.Enabled = True
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

