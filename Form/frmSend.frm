VERSION 5.00
Begin VB.Form frmSend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kirim Pesan"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2805
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSend 
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton cmdSend 
         Caption         =   "Kirim"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   720
         MaxLength       =   3
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         MaxLength       =   3
         TabIndex        =   0
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SettingAwal
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSend = Nothing
End Sub

Private Sub cmdSend_Click()
    If Not CekValid Then Exit Sub
    
    Dim strIP(3) As String
    
    strIP(0) = Me.txtIP(0).Text
    strIP(1) = Me.txtIP(1).Text
    strIP(2) = Me.txtIP(2).Text
    strIP(3) = Me.txtIP(3).Text
    
    frmMain.KirimData strIP
    
    Unload Me
End Sub

Private Function CekValid() As Boolean
    CekValid = True
    
    Dim intCounter As Integer
    
    For intCounter = 0 To Me.txtIP.Count - 1
        If Trim(Me.txtIP(intCounter).Text) = "" Then
            Me.txtIP(intCounter).SetFocus
            
            CekValid = False
            
            MsgBox "IP Ke " & intCounter + 1 & " Kosong", vbCritical, Me.Caption
            
            Exit For
        ElseIf Not IsNumeric(Me.txtIP(intCounter).Text) Then
            Me.txtIP(intCounter).SetFocus
            
            CekValid = False
            
            MsgBox "IP Ke " & intCounter + 1 & " Bukan Angka", vbCritical, Me.Caption
            
            Exit For
        End If
    Next intCounter
End Function

Private Sub SettingAwal()
    Dim strIPSendiri() As String
    
    strIPSendiri = frmMain.IPSendiri
    
    Me.txtIP(0).Text = strIPSendiri(0)
    Me.txtIP(1).Text = strIPSendiri(1)
    Me.txtIP(2).Text = strIPSendiri(2)
    Me.txtIP(3).Text = strIPSendiri(3)
End Sub
