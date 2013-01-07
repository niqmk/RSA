VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin TabDlg.SSTab sstHelp 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Contents"
      TabPicture(0)   =   "frmHelp.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblContents(4)"
      Tab(0).Control(1)=   "lblContents(3)"
      Tab(0).Control(2)=   "lblContents(2)"
      Tab(0).Control(3)=   "lblContents(1)"
      Tab(0).Control(4)=   "lblContents(0)"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Login"
      TabPicture(1)   =   "frmHelp.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblLogin(4)"
      Tab(1).Control(1)=   "lblLogin(3)"
      Tab(1).Control(2)=   "lblLogin(2)"
      Tab(1).Control(3)=   "lblLogin(1)"
      Tab(1).Control(4)=   "lblLogin(0)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Enkripsi"
      TabPicture(2)   =   "frmHelp.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblEnkripsi(3)"
      Tab(2).Control(1)=   "lblEnkripsi(2)"
      Tab(2).Control(2)=   "lblEnkripsi(1)"
      Tab(2).Control(3)=   "lblEnkripsi(0)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Dekripsi"
      TabPicture(3)   =   "frmHelp.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblDekripsi(0)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblDekripsi(1)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblDekripsi(2)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblDekripsi(3)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin VB.Label lblDekripsi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Tombol Dekripsi"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   1560
         Width           =   1320
      End
      Begin VB.Label lblDekripsi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Masukkan Kunci Pertama dan Kunci Kedua"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   3285
      End
      Begin VB.Label lblDekripsi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Tulis Pesan"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label lblDekripsi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manual Pemakaian Program Dekripsi"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   2610
      End
      Begin VB.Label lblEnkripsi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Tombol Enkripsi"
         Height          =   195
         Index           =   3
         Left            =   -74760
         TabIndex        =   15
         Top             =   1560
         Width           =   1305
      End
      Begin VB.Label lblEnkripsi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Masukkan Kunci Pertama dan Kunci Kedua"
         Height          =   195
         Index           =   2
         Left            =   -74760
         TabIndex        =   14
         Top             =   1200
         Width           =   3285
      End
      Begin VB.Label lblEnkripsi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Tulis Pesan"
         Height          =   195
         Index           =   1
         Left            =   -74760
         TabIndex        =   13
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label lblEnkripsi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manual Pemakaian Program Enkripsi"
         Height          =   195
         Index           =   0
         Left            =   -74760
         TabIndex        =   12
         Top             =   480
         Width           =   2595
      End
      Begin VB.Label lblLogin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. Tombol Batal"
         Height          =   195
         Index           =   4
         Left            =   -74760
         TabIndex        =   11
         Top             =   1920
         Width           =   1110
      End
      Begin VB.Label lblLogin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Tombol OK"
         Height          =   195
         Index           =   3
         Left            =   -74760
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblLogin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Password"
         Height          =   195
         Index           =   2
         Left            =   -74760
         TabIndex        =   9
         Top             =   1200
         Width           =   870
      End
      Begin VB.Label lblLogin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Nama Pemakai"
         Height          =   195
         Index           =   1
         Left            =   -74760
         TabIndex        =   8
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label lblLogin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manual Pemakaian Program Login"
         Height          =   195
         Index           =   0
         Left            =   -74760
         TabIndex        =   7
         Top             =   480
         Width           =   2430
      End
      Begin VB.Label lblContents 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4. Dekripsi"
         Height          =   195
         Index           =   4
         Left            =   -74760
         TabIndex        =   6
         Top             =   1920
         Width           =   750
      End
      Begin VB.Label lblContents 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3. Enkripsi"
         Height          =   195
         Index           =   3
         Left            =   -74760
         TabIndex        =   5
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblContents 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. Login"
         Height          =   195
         Index           =   2
         Left            =   -74760
         TabIndex        =   4
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label lblContents 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. Panduan Halaman"
         Height          =   195
         Index           =   1
         Left            =   -74760
         TabIndex        =   3
         Top             =   840
         Width           =   1500
      End
      Begin VB.Label lblContents 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manual Pemakaian Program"
         Height          =   195
         Index           =   0
         Left            =   -74760
         TabIndex        =   2
         Top             =   480
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.sstHelp.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmHelp = Nothing
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
