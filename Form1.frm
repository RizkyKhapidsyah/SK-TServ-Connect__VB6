VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{FC7C887E-70BD-4ADB-8BED-8681D74F36D1}#1.0#0"; "msrdp.ocx"
Begin VB.Form Form1 
   Caption         =   "TSERV Connect!"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   13335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin MSTSCLibCtl.MsTscAx MsTscAx1 
      Height          =   7815
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   13095
      Server          =   ""
      Domain          =   ""
      UserName        =   ""
      FullScreen      =   ""
      StartConnected  =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   8685
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "TSERV Connect!"
            TextSave        =   "TSERV Connect!"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "18:12"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "19/04/2021"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   35278
            MinWidth        =   35278
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command4 
      Caption         =   "About"
      Height          =   375
      Left            =   12000
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Server:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

On Error GoTo Handle:
MsTscAx1.Server = Text1.Text
MsTscAx1.Connect
Exit Sub

Handle:
MsgBox ("Error." & vbCrLf & _
"Number: " & Err.Number & _
"Description: " & Err.Description & _
"Source: " & Err.Source)
Err.Clear
Resume Next

End Sub

Private Sub Command2_Click()
MsTscAx1.Disconnect
End Sub

Private Sub Command4_Click()
Load frmAbout
Me.Hide
frmAbout.Show
End Sub

Private Sub Form_Load()
MsTscAx1.ConnectingText = "Connecting to " & Text1.Text & "..."
MsTscAx1.DisconnectedText = "Disconnected."
End Sub

Private Sub Form_Resize()
MsTscAx1.Width = Form1.Width - 1000
MsTscAx1.Left = 500
MsTscAx1.Height = Form1.Height - 2000
Command4.Left = Form1.Width - 1700

End Sub

