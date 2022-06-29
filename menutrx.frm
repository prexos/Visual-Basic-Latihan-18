VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form menutrx 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8895
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
   ScaleHeight     =   5010
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton back 
      Caption         =   "Menu Utama"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6240
      TabIndex        =   22
      Top             =   9240
      Width           =   2055
   End
   Begin VB.Data Data2 
      Caption         =   "Data Transaksi"
      Connect         =   "Access"
      DatabaseName    =   "D:\Kuliah\Semester 4\Pemrograman\Visual-Basic-Latihan-18-main\database18.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "T_TRX"
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Data Data1 
      Caption         =   "Data Nasabah"
      Connect         =   "Access"
      DatabaseName    =   "D:\Kuliah\Semester 4\Pemrograman\Visual-Basic-Latihan-18-main\database18.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "T_NASABAH"
      Top             =   3840
      Width           =   2775
   End
   Begin VB.CommandButton end 
      Caption         =   "Selesai"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7560
      TabIndex        =   18
      Top             =   7800
      Width           =   2295
   End
   Begin VB.CommandButton Save 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      TabIndex        =   17
      Top             =   7800
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   8640
      TabIndex        =   16
      Top             =   6480
      Width           =   2175
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   6000
      TabIndex        =   15
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Height          =   360
      Left            =   3480
      TabIndex        =   14
      Top             =   6480
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   9240
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   118620161
      CurrentDate     =   44741
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   2160
      TabIndex        =   4
      Top             =   2520
      Width           =   4335
      Begin VB.TextBox Text5 
         Height          =   360
         Left            =   1440
         TabIndex        =   11
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Kota"
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Alamat"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Nama"
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Nomor Nasabah"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   3240
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Jl. Dayang Sumbi 07. Bandung Jabar"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   23
      Top             =   720
      Width           =   8175
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Sisa Hutang"
      Height          =   375
      Left            =   8640
      TabIndex        =   21
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Dibayar"
      Height          =   375
      Left            =   6000
      TabIndex        =   20
      Top             =   6120
      Width           =   2295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Total Hutang"
      Height          =   255
      Left            =   3480
      TabIndex        =   19
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   8280
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Nomor"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PT. SANGKURIANG"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   7935
   End
End
Attribute VB_Name = "menutrx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub back_Click()
    menuutama.Show
End Sub

Private Sub end_Click()
    End
End Sub

Private Sub Form_Load()
    menutrx.WindowState = 2
End Sub

Private Sub Save_Click()
    'Add data
    Data2.Recordset.AddNew
    Data2.Recordset!nomortrx = Text1.Text
    Data2.Recordset!tanggal = DTPicker1.Value
    Data2.Recordset!nokodenasabah = Text2.Text
    Data2.Recordset!jmlpembayaran = Text7.Text
    Data2.Recordset.Update
    'Additional
    Respon = MsgBox("Data berhasil disimpan! Ingin keluar?", vbYesNo, "Berhasil")
    If Respon = vbNo Then
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""
        Text7.Text = ""
        Text8.Text = ""
    Else
        End
    End If
End Sub

Private Sub Text2_LostFocus()
    Cari = "nomorkode='" + Text2.Text + "'"
    Data1.Recordset.FindFirst Cari
    If Data1.Recordset.NoMatch Then
        Respon = MsgBox("Data Nasabah tidak Ditemukan, cari lainnya?", vbYesNo, "Peringatan")
        If Respon = vbYes Then
            Text2.Text = ""
            Text2.SetFocus
        End If
    Else
        Text3.Text = Data1.Recordset!nama
        Text4.Text = Data1.Recordset!alamat
        Text5.Text = Data1.Recordset!kota
        Text6.Text = Data1.Recordset!totalhutang
        Text3.Enabled = False
        Text4.Enabled = False
        Text5.Enabled = False
        Text6.Enabled = False
    End If
End Sub

Private Sub Text6_Change()
    Text8.Text = Val(Text6.Text) - Val(Text7.Text)
End Sub

Private Sub Text7_Change()
    Text8.Text = Val(Text6.Text) - Val(Text7.Text)
End Sub

Private Sub Text8_Change()
    Text8.Enabled = False
End Sub
