VERSION 5.00
Begin VB.Form laporan 
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Menu Utama"
      Height          =   1095
      Left            =   4560
      TabIndex        =   2
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Laporan Penerimaan Kas"
      Height          =   1335
      Left            =   5880
      TabIndex        =   1
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Laporan Daftar Nasabah"
      Height          =   1335
      Left            =   2880
      TabIndex        =   0
      Top             =   2280
      Width           =   2535
   End
End
Attribute VB_Name = "Laporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    LapDaftarNasabah.Show
End Sub

Private Sub Command2_Click()
    LapPenerimaanKas.Show
End Sub

Private Sub Command3_Click()
    menuutama.Show
End Sub

Private Sub Form_Load()
    Laporan.WindowState = 2
End Sub
