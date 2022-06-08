VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
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
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Delete Data"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7320
      TabIndex        =   17
      Top             =   8040
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search Data"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9960
      TabIndex        =   16
      Top             =   8040
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Data"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      TabIndex        =   15
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton exit 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7320
      TabIndex        =   13
      Top             =   9480
      Width           =   2295
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   2535
      Left            =   3360
      OleObjectBlob   =   "Form1.frx":0014
      TabIndex        =   12
      Top             =   4200
      Width           =   9375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Kuliah\Semester 4\Pemrograman\Visual-Basic-Latihan-10-main\penjualan.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TCUSTOMER"
      Top             =   6960
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      DataField       =   "TELPONCUST"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   10200
      TabIndex        =   11
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      DataField       =   "KONTAKCUST"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   10200
      TabIndex        =   10
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      DataField       =   "KOTACUST"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   10200
      TabIndex        =   9
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      DataField       =   "ALAMATCUST"
      DataSource      =   "Data1"
      Height          =   615
      Left            =   4320
      TabIndex        =   4
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      DataField       =   "NAMACUST"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   4320
      TabIndex        =   3
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      DataField       =   "NOCUST"
      DataSource      =   "Data1"
      Height          =   360
      Left            =   4320
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "PENGELOLAAN DATA CUSTOMER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      TabIndex        =   14
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label6 
      Caption         =   "No. Telp"
      Height          =   255
      Left            =   9240
      TabIndex        =   8
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Kontak"
      Height          =   255
      Left            =   9240
      TabIndex        =   7
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Kota"
      Height          =   255
      Left            =   9240
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Alamat"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Nama"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Nomor"
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   1920
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'menambahkan data
    If Text1.Text <> "" Then
    Data1.Recordset.AddNew
    End If
    Text1.SetFocus
End Sub

Private Sub Command2_Click()
    'mencari data
    Respon = vbYes
    While Respon = vbYes
        Respon = InputBox("Masukkan Nomor Customer", "Input Data Pencarian")
        Cari = "NOCUST='" + Respon + "'"
        Data1.Recordset.FindFirst Cari
        If Data1.Recordset.NoMatch Then
            Respon = MsgBox("Data Dicari Tidak Ada! Cari Lainnya?", vbYesNo, "Cari Data")
        Else
            Respon = vbNo
        End If
    Wend
End Sub

Private Sub Command3_Click()
    'hapus data
    Respon = MsgBox("Menghapus Data Record Ini?", vbYesNo, "Alert!")
    If Respon = vbYes Then
        Data1.Recordset.Delete
        Data1.Refresh
    End If
    Text1.SetFocus
End Sub

Private Sub exit_Click()
    End
End Sub

Private Sub Form_Activate()
    'Window
    Form1.WindowState = 2
    
    'perinth agar cek terhadap record terakhir, jika tidak kosong buat satu record kosong
    'baru ini dimaksud supaya tampilan form pertama adalah kosong
    Data1.Recordset.MoveLast
    If Data1.Recordset!NOCUST <> "" Then
        Data1.Recordset.AddNew
    End If
End Sub
