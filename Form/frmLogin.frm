VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MENU LOGIN"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMasuk 
      Caption         =   "Masuk"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Frame fraMain 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtPass_Pemakai 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtNama_Pemakai 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lblPass_Pemakai 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblNama_Pemakai 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pemakai"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private conApp As ADODB.Connection

Private Sub Form_Load()
    SettingAwal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    conApp.Close
    
    Set conApp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmLogin = Nothing
End Sub

Private Sub cmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdMasuk_Click()
    If Not CekPemakai Then Exit Sub
    
    frmMain.Show
    
    Unload Me
End Sub

Private Sub SettingAwal()
    Set conApp = New ADODB.Connection
    conApp.CursorLocation = adUseClient
    conApp.Provider = "Microsoft.Jet.OLEDB.4.0"
    conApp.Open mdlGlobal.strPath & "RSA.mdb"
End Sub

Private Function CekPemakai() As Boolean
    If Trim(Me.txtNama_Pemakai.Text) = "" Then
        MsgBox "Nama Pemakai Kosong", vbCritical, Me.Caption
        
        Me.txtNama_Pemakai.SetFocus
        
        CekPemakai = False
        
        Exit Function
    ElseIf Trim(Me.txtPass_Pemakai.Text) = "" Then
        MsgBox "Password Kosong", vbCritical, Me.Caption
        
        Me.txtPass_Pemakai.SetFocus
        
        CekPemakai = False
        
        Exit Function
    End If
    
    Dim rstPemakai As New ADODB.Recordset
    rstPemakai.CursorLocation = adUseClient
    rstPemakai.Open "SELECT * FROM Pemakai WHERE Nama_Pemakai='" & Trim(Me.txtNama_Pemakai.Text) & "' AND Pass_Pemakai='" & Trim(Me.txtPass_Pemakai.Text) & "'", conApp, adOpenDynamic, adLockOptimistic
    
    If rstPemakai.RecordCount > 0 Then
        CekPemakai = True
    Else
        MsgBox "Nama Pemakai atau Password Tidak Benar", vbCritical, "Login"
    
        CekPemakai = False
    End If
    
    rstPemakai.Close
    Set rstPemakai = Nothing
End Function
