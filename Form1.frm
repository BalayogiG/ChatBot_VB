VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "chatbot"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Height          =   5715
      ItemData        =   "Form1.frx":0000
      Left            =   9720
      List            =   "Form1.frx":0002
      TabIndex        =   5
      Top             =   120
      Width           =   3615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   360
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Jarvis\Desktop\CHATBOT\CHATCOMMANDS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Jarvis\Desktop\CHATBOT\CHATCOMMANDS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from commands"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEND"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   1695
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3240
      Width           =   6975
   End
   Begin VB.TextBox Text1 
      Height          =   1695
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1080
      Width           =   8535
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   2040
      Picture         =   "Form1.frx":0004
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "CHAT BOT"
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "MESSAGE :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Menu mnucommands 
      Caption         =   "commands"
   End
   Begin VB.Menu muexit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset


Private Sub Command1_Click()

Dim strcommand As String
Dim strresponse As String

strcommand = Text2.Text
Adodc1.RecordSource = "select RESPONSES from commands where COMMANDS='" + Text2.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
strresponse = "I didn't get you"
Text1.Text = "USER: " + Text2.Text & vbCrLf & "SYSTEM: " + strresponse
Else
strresponse = Adodc1.Recordset.Fields("RESPONSES")
Text1.Text = "USER: " + Text2.Text & vbCrLf & "SYSTEM: " + strresponse
End If

                        
If strcommand = "what is the time now?" Then
strresponse = Time
Text1.Text = "USER: " + Text2.Text & vbCrLf & "SYSTEM: " + strresponse
End If

If strcommand = "open internet" Then
Shell "C:\Program Files (x86)\Internet Explorer\iexplore.exe", vbNormalFocus
End If

If strcommand = "open notepad" Then
Shell "C:\Windows\notepad.exe", vbNormalFocus
End If

If strcommand = "show commands" Then
    If Me.Width = 9780 Then
        Me.Width = 13500
    Else
        Me.Width = 9780
    End If
End If

Set objspeech = CreateObject("SAPI.spvoice")
objspeech.Speak strresponse

End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""

Dim x As Integer
Dim y As Integer

Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset

cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.ConnectionString = App.Path & "\CHATCOMMANDS.mdb"
cn.Open

rs.Source = "select COMMANDS from commands"
rs.Open , cn, adOpenKeyset, adLockOptimistic

Do Until rs.EOF

List1.AddItem rs!COMMANDS

rs.MoveNext
Loop
rs.Close
cn.Close

Set rs = Nothing
Set cn = Nothing

End Sub

Private Sub mnucommands_Click()
If Me.Width = 9780 Then
Me.Width = 13500
Else
Me.Width = 9780
End If
End Sub

Private Sub muexit_Click()
End
End Sub
