VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Hanya Huruf Kecil Boleh Dientri ke TextBox"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   1440
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'Created by Rizky Khapidsyah
'Source code dimulai dari sini

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
End Sub

Private Sub Text1_Change()  'Text1 menggunakan event
                            'Change
Dim posisi As Integer
posisi = Text1.SelStart
  Text1.Text = LCase(Text1.Text)
  Text1.SelStart = posisi
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)  'Text2 'menggunakan KeyPress
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub

