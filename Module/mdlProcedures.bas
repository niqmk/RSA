Attribute VB_Name = "mdlProcedures"
Option Explicit

Public Function IsPrime(ByVal intValue As Integer) As Boolean
    Dim blnValid As Boolean
    
    blnValid = False
    
    Dim intCount As Integer
    
    intCount = 0
    
    If intValue = 1 Or intValue = 2 Then intCount = intCount + 1

    Dim intCounter As Integer
    
    For intCounter = 1 To intValue
        If (intValue Mod intCounter) = 0 Then intCount = intCount + 1
    Next intCounter
    
    If Not intCount > 2 Then blnValid = True
    
    IsPrime = blnValid
End Function

Public Function IsRound(ByVal curValue As Currency) As Boolean
    Dim strValue As String
    
    strValue = CStr(curValue)
    
    If InStr(strValue, ".") > 0 Then
        curValue = curValue - (CCur(Left(strValue, InStr(strValue, "."))))
    Else
        IsRound = False
        
        Exit Function
    End If
    
    If curValue * 10000 > 0 Then
        IsRound = True
    Else
        IsRound = False
    End If
End Function

Public Function GetStringToSequence(ByVal strValue As String) As Integer
    Select Case UCase(strValue)
        Case "A"
            GetStringToSequence = 1
        Case "B"
            GetStringToSequence = 2
        Case "C"
            GetStringToSequence = 3
        Case "D"
            GetStringToSequence = 4
        Case "E"
            GetStringToSequence = 5
        Case "F"
            GetStringToSequence = 6
        Case "G"
            GetStringToSequence = 7
        Case "H"
            GetStringToSequence = 8
        Case "I"
            GetStringToSequence = 9
        Case "J"
            GetStringToSequence = 10
        Case "K"
            GetStringToSequence = 11
        Case "L"
            GetStringToSequence = 12
        Case "M"
            GetStringToSequence = 13
        Case "N"
            GetStringToSequence = 14
        Case "O"
            GetStringToSequence = 15
        Case "P"
            GetStringToSequence = 16
        Case "Q"
            GetStringToSequence = 17
        Case "R"
            GetStringToSequence = 18
        Case "S"
            GetStringToSequence = 19
        Case "T"
            GetStringToSequence = 20
        Case "U"
            GetStringToSequence = 21
        Case "V"
            GetStringToSequence = 22
        Case "W"
            GetStringToSequence = 23
        Case "X"
            GetStringToSequence = 24
        Case "Y"
            GetStringToSequence = 25
        Case "Z"
            GetStringToSequence = 26
    End Select
End Function

Public Function GetSequenceToString(ByVal strSequence As String) As String
    Select Case strSequence
        Case 1
            GetSequenceToString = "A"
        Case 2
            GetSequenceToString = "B"
        Case 3
            GetSequenceToString = "C"
        Case 4
            GetSequenceToString = "D"
        Case 5
            GetSequenceToString = "E"
        Case 6
            GetSequenceToString = "F"
        Case 7
            GetSequenceToString = "G"
        Case 8
            GetSequenceToString = "H"
        Case 9
            GetSequenceToString = "I"
        Case 10
            GetSequenceToString = "J"
        Case 11
            GetSequenceToString = "K"
        Case 12
            GetSequenceToString = "L"
        Case 13
            GetSequenceToString = "M"
        Case 14
            GetSequenceToString = "N"
        Case 15
            GetSequenceToString = "O"
        Case 16
            GetSequenceToString = "P"
        Case 17
            GetSequenceToString = "Q"
        Case 18
            GetSequenceToString = "R"
        Case 19
            GetSequenceToString = "S"
        Case 20
            GetSequenceToString = "T"
        Case 21
            GetSequenceToString = "U"
        Case 22
            GetSequenceToString = "V"
        Case 23
            GetSequenceToString = "W"
        Case 24
            GetSequenceToString = "X"
        Case 25
            GetSequenceToString = "Y"
        Case 26
            GetSequenceToString = "Z"
        Case Else
            GetSequenceToString = " "
    End Select
End Function

Public Function Modulation(ByVal strSource As String, ByVal intDivision As Integer) As String
    Dim strResult As String
    
    If IsRoundPlusOne(strSource, intDivision) Then
        strResult = Round(strSource / intDivision, 0) - 1
    Else
        strResult = Round(strSource / intDivision, 0)
    End If
    
    Dim strDivision As String
    
    strDivision = (strResult * intDivision)
    
    strResult = strSource - strDivision
    
    Modulation = strResult
End Function

Private Function IsRoundPlusOne(ByVal strSource As String, ByVal intDivision As String) As Boolean
    Dim strValue As String
    
    strValue = CStr(strSource / intDivision)
    
    If InStr(strValue, ".") > 0 Then
        strValue = Mid(strValue, InStr(strValue, ".") + 1)
        
        If CInt(Left(strValue, 1)) >= 5 Then
            IsRoundPlusOne = True
        Else
            IsRoundPlusOne = False
        End If
    Else
        IsRoundPlusOne = False
    End If
End Function
