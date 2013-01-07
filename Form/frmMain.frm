VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PENGAMANAN DATA DENGAN ALGORITMA RSA"
   ClientHeight    =   4095
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   1535
      ButtonWidth     =   1244
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "imlMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save As"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Connect"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Send"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList imlMain 
         Left            =   6600
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2452
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":48A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6CF6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":73B3
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":9805
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraPesan 
      Caption         =   "Pesan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   7095
      Begin MSWinsockLib.Winsock wskNetwork 
         Index           =   0
         Left            =   5040
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.TextBox txtPesan 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   4695
      End
      Begin MSComDlg.CommonDialog cdlFile 
         Left            =   6480
         Top             =   2400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdDekripsi 
         Caption         =   "Dekripsi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtKunci_Pertama 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtKunci_Kedua 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         MaxLength       =   4
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdEnkripsi 
         Caption         =   "Enkripsi"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   2520
         Width           =   1335
      End
      Begin MSWinsockLib.Winsock wskNetwork 
         Index           =   1
         Left            =   5520
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label lblKunci_Pertama 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kunci Pertama"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4920
         TabIndex        =   5
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label lblKunci_Kedua 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kunci Kedua"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4920
         TabIndex        =   6
         Top             =   720
         Width           =   1110
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save As"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SettingAwal
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mdlGlobal.fso = Nothing
    
    Me.wskNetwork(0).Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMain = Nothing
End Sub

Private Sub mnuNew_Click()
    PilihanNew
End Sub

Private Sub mnuOpen_Click()
    PilihanOpen
End Sub

Private Sub mnuSave_Click()
    PilihanSave
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuHelp_Click()
    frmHelp.Show vbModal
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1:
            PilihanNew
        Case 2:
            PilihanOpen
        Case 3:
            PilihanSave
        Case 4:
            PilihanKoneksi
        Case 5:
            PilihanSend
        Case 6:
            PilihanHelp
    End Select
End Sub

Private Sub txtKunci_Pertama_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtKunci_Pertama.Text) Then Me.txtKunci_Pertama.Text = "3"
    
    If Not CInt(Me.txtKunci_Pertama.Text) > 2 Then Me.txtKunci_Pertama.Text = "3"
End Sub

Private Sub txtKunci_Kedua_Validate(Cancel As Boolean)
    If Not IsNumeric(Me.txtKunci_Kedua.Text) Then Me.txtKunci_Kedua.Text = "3"
    
    If Not CInt(Me.txtKunci_Kedua.Text) > 2 Then Me.txtKunci_Kedua.Text = "3"
End Sub

Private Sub cmdEnkripsi_Click()
    If Not CekValid Then Exit Sub
    
    Me.txtPesan.Text = mdlRSA.StartEncrypt(Me.txtPesan.Text, CInt(Me.txtKunci_Pertama.Text), CInt(Me.txtKunci_Kedua.Text))
End Sub

Private Sub cmdDekripsi_Click()
    If Not CekValid Then Exit Sub
    
    Me.txtPesan.Text = mdlRSA.StartDecrypt(Me.txtPesan.Text, CInt(Me.txtKunci_Pertama.Text), CInt(Me.txtKunci_Kedua.Text))
End Sub

Private Sub SettingAwal()
    Set mdlGlobal.fso = New FileSystemObject

    Me.cdlFile.Filter = "File RSA (*.txt)|*.txt"
    Me.cdlFile.InitDir = mdlGlobal.strPath
    Me.cdlFile.CancelError = True
    
    Me.wskNetwork(0).LocalPort = mdlGlobal.lngPort
    Me.wskNetwork(0).Listen

    If Me.wskNetwork(0).State = sckListening Then
        Me.Caption = Me.Caption & " ( " & Me.wskNetwork(0).LocalIP & " )"
    End If
End Sub

Private Function CekValid() As Boolean
    Dim blnValid As Boolean
    
    blnValid = True
    
    If Not IsNumeric(Me.txtKunci_Pertama.Text) Then Me.txtKunci_Pertama.Text = "3"
    
    If Not IsNumeric(Me.txtKunci_Kedua.Text) Then Me.txtKunci_Kedua.Text = "3"
    
    If Trim(Me.txtPesan.Text) = "" Then
        MsgBox "Pesan Kosong", vbExclamation + vbOKOnly, Me.Caption
        
        Me.txtPesan.SetFocus
        
        blnValid = False
    ElseIf Not mdlProcedures.IsPrime(CInt(Me.txtKunci_Pertama.Text)) Then
        MsgBox "Kunci Pertama Bukan Bilangan Prima", vbExclamation + vbOKOnly, Me.Caption
        
        Me.txtKunci_Pertama.SetFocus
    
        blnValid = False
    ElseIf Not mdlProcedures.IsPrime(CInt(Me.txtKunci_Kedua.Text)) Then
        MsgBox "Kunci Kedua Bukan Bilangan Prima", vbExclamation + vbOKOnly, Me.Caption
        
        Me.txtKunci_Kedua.SetFocus
    
        blnValid = False
    End If
    
    CekValid = blnValid
End Function

Private Sub PilihanNew()
    If MsgBox("Apakah Anda Ingin Hapus Semua", vbInformation + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Me.txtPesan.Text = ""
    Me.txtKunci_Pertama.Text = ""
    Me.txtKunci_Kedua.Text = ""
End Sub

Private Sub PilihanOpen()
    On Local Error GoTo ErrHandler

    With Me.cdlFile
        .ShowOpen
        
        If Trim(.FileName) = "" Then Exit Sub
        
        Dim txsFile As TextStream
        
        Set txsFile = mdlGlobal.fso.OpenTextFile(.FileName, ForReading)
        
        Me.txtPesan.Text = txsFile.ReadAll
        
        txsFile.Close
        Set txsFile = Nothing
    End With
    
    Exit Sub
    
ErrHandler:
End Sub

Private Sub PilihanSave()
    On Local Error GoTo ErrHandler

    With Me.cdlFile
        .ShowSave
        
        If Trim(.FileName) = "" Then Exit Sub
        
        Dim txsFile As TextStream
        
        Set txsFile = mdlGlobal.fso.CreateTextFile(.FileName, True)
        
        txsFile.Write Me.txtPesan.Text
        
        txsFile.Close
        Set txsFile = Nothing
        
        MsgBox "Pesan Sudah Disimpan di" & vbCrLf & .FileName, vbInformation + vbOKOnly, Me.Caption
    End With
    
    Exit Sub
    
ErrHandler:
End Sub

Private Sub PilihanKoneksi()
    frmSend.Show vbModal
End Sub

Private Sub PilihanSend()
    If Me.wskNetwork(1).State = sckClosed Then
        MsgBox "Anda Belum Terkoneksi IP Yang Ingin Dituju", vbCritical, Me.Caption
        
        Exit Sub
    End If
    
    If Me.wskNetwork(1).State = sckConnected Then
        Me.wskNetwork(1).SendData Me.txtPesan.Text
    End If
End Sub

Private Sub PilihanHelp()
    frmHelp.Show vbModal
End Sub

Public Sub KirimData(ByRef strIP() As String)
    If Not Me.wskNetwork(1).State = sckClosed Then
        Me.wskNetwork(1).SendData "--->><<---"
        
        DoEvents
        
        Me.wskNetwork(1).Close
    End If
    
    Me.wskNetwork(1).RemotePort = mdlGlobal.lngPort
    Me.wskNetwork(1).RemoteHost = strIP(0) & "." & strIP(1) & "." & strIP(2) & "." & strIP(3)
    
    Me.wskNetwork(1).Connect
End Sub

Public Property Get IPSendiri() As String()
    Dim strIP() As String
    
    strIP = Split(Me.wskNetwork(0).LocalIP, ".")
    
    IPSendiri = strIP
End Property

Public Property Get IPLain() As String
    IPLain = Me.wskNetwork(0).RemoteHostIP
End Property

Private Sub wskNetwork_Close(Index As Integer)
    Select Case Index
        Case 0:
            Me.wskNetwork(0).Close
            Me.wskNetwork(0).LocalPort = mdlGlobal.lngPort
            Me.wskNetwork(0).Listen
    End Select
End Sub

Private Sub wskNetwork_Connect(Index As Integer)
    Me.fraPesan.Caption = "Pesan" & " Kirim Ke " & Me.wskNetwork(1).RemoteHostIP
End Sub

Private Sub wskNetwork_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If Not Me.wskNetwork(0).State = sckClosed Then
        Me.wskNetwork(0).Close
    End If
    
    Me.wskNetwork(0).Accept requestID
End Sub

Private Sub wskNetwork_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strText As String
    
    Me.wskNetwork(0).GetData strText, vbString
    
    If Trim(strText) = "--->><<---" Then
        Me.wskNetwork(0).Close
        
        Me.wskNetwork(0).LocalPort = mdlGlobal.lngPort
        Me.wskNetwork(0).Listen
    Else
        Me.txtPesan.Text = strText
    End If
End Sub

Private Sub wskNetwork_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbCritical, Number
    
    If Not Me.wskNetwork(1).State = sckClosed Then
        Me.wskNetwork(1).Close
    End If
End Sub
