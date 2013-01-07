Attribute VB_Name = "mdlMain"
Option Explicit

Public Sub Main()
    SettingVariable
    
    frmLogin.Show
End Sub

Private Sub SettingVariable()
    mdlGlobal.strPath = App.Path
    
    If Not Right(mdlGlobal.strPath, 1) = "\" Then mdlGlobal.strPath = mdlGlobal.strPath & "\"
End Sub
