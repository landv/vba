VERSION 5.00
Object = "{CC802D05-AE07-4C15-B496-DB9D22AA0A84}#1.0#0"; "rdpencom.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17265
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   17265
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Button_test 
      Caption         =   "Button_test"
      Height          =   795
      Left            =   14070
      TabIndex        =   2
      Top             =   150
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Height          =   8985
      Left            =   9720
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   150
      Width           =   4095
   End
   Begin RDPCOMAPILibCtl.RDPViewer RDPViewer1 
      Height          =   9255
      Left            =   90
      OleObjectBlob   =   "Form1.frx":0006
      TabIndex        =   0
      Top             =   90
      Width           =   9465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Button_test_Click()
    RDPViewer1.SmartSizing = True
    RDPViewer1.Connect(Text1.Text, "administrator", "1234")
    RDPViewer1.RequestControl (CTRL_LEVEL_MAX)

End Sub

Private Sub Form_Load()
    Dim r As New RDPSession
    Debug.Print Screen.Width / Screen.TwipsPerPixelX & "," & Screen.Height / Screen.TwipsPerPixelY
    r.Open
    Dim rdpinv As RDPSRAPIInvitation
    Set rdpinv = r.Invitations.CreateInvitation("baseAuth", "groupName", "1234", 64)
    Debug.Print rdpinv.ConnectionString
    'r.Close
    Text1.Text = rdpinv.ConnectionString
    Main.Text1.Text = Main.Text1.Text + vbCrLf + rdpinv.ConnectionString
End Sub
