Attribute VB_Name = "mdlRSA"
Option Explicit

Public Function StartEncrypt(ByVal strPlainText As String, ByVal intP As Integer, ByVal intQ As Integer) As String
    Dim intM As Integer
    Dim intN As Integer
    
    GetMandN intM, intN, intP, intQ
    
    Dim intE As Integer
    
    intE = GetKeyForEncrypt(intM)
    
    StartEncrypt = EncryptIt(strPlainText, intE, intN)
End Function

Public Function StartDecrypt(ByVal strEncryptText As String, ByVal intP As Integer, ByVal intQ As Integer) As String
    Dim intM As Integer
    Dim intN As Integer
    
    GetMandN intM, intN, intP, intQ
    
    Dim intE As Integer
    
    intE = GetKeyForEncrypt(intM)
    
    Dim intD As Integer
    
    intD = GetKeyForDecrypt(intM, intN, intE)
    
    StartDecrypt = DecryptIt(strEncryptText, intD, intN)
End Function

Private Sub GetMandN(ByRef intM As Integer, ByRef intN As Integer, ByVal intP As Integer, ByVal intQ As Integer)
    intM = (intP - 1) * (intQ - 1)
    intN = intP * intQ
End Sub

Private Function GetKeyForEncrypt(ByVal intM As Integer) As Integer
    Dim intCounter As Integer
    
    For intCounter = 2 To (intM - 1)
        If Not (intM Mod intCounter) = 0 Then
            GetKeyForEncrypt = intCounter
            
            Exit For
        End If
    Next intCounter
End Function

Private Function GetKeyForDecrypt(ByVal intM As Integer, ByVal intN As Integer, ByVal intE As Integer) As Integer
    Dim dblValue As Double
    
    Dim intD As Integer
    
    intD = 0
    
    Dim intCounter As Integer

    For intCounter = 0 To (intN - 1)
        dblValue = (1 + (intCounter * intM)) / intE

        If Not mdlProcedures.IsRound(CCur(dblValue)) Then
            intD = CInt(dblValue)

            Exit For
        End If
    Next intCounter
    
    GetKeyForDecrypt = intD
End Function

Private Function EncryptIt(ByVal strText As String, ByVal intE As Integer, ByVal intN As Integer) As String
    Dim strEncrypt As String
    
    strEncrypt = ""
    
    Dim intSequence As Integer
    
    Dim strC As Integer
    
    Dim intCounter As Integer
    
    For intCounter = 1 To Len(strText)
        If Trim(Mid(strText, intCounter, 1)) = "" Then
            strEncrypt = strEncrypt & "  | "
        Else
            intSequence = mdlProcedures.GetStringToSequence(Mid(strText, intCounter, 1))
            
            strC = mdlProcedures.Modulation((intSequence ^ intE), intN)
            
            strEncrypt = strEncrypt & CStr(strC) & " | "
        End If
    Next intCounter
    
    If Not Trim(strEncrypt) = "" Then
        strEncrypt = Left(strEncrypt, Len(strEncrypt) - 3)
    End If
    
    EncryptIt = strEncrypt
End Function

Private Function DecryptIt(ByVal strText As String, ByVal intD As Integer, ByVal intN As Integer) As String
    Dim strDecrypt As String
    
    strDecrypt = ""
    
    Dim strSequence As String
    
    Dim strC As String
    
    Dim strTemp As String
    
    Dim intCounter As Integer
    
    Dim strTextTemp() As String
    
    strTextTemp = Split(strText, " | ")
    
    For intCounter = 0 To UBound(strTextTemp)
        If Trim(strTextTemp(intCounter)) = "" Then
            strDecrypt = strDecrypt & " "
        Else
            strSequence = strTextTemp(intCounter)
            
            strTemp = strSequence ^ CLng(intD)
            
            strC = mdlProcedures.Modulation(strTemp, intN)
            
            strDecrypt = strDecrypt & mdlProcedures.GetSequenceToString(strC)
        End If
    Next intCounter
    
    DecryptIt = strDecrypt
End Function
