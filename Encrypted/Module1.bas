Attribute VB_Name = "Module1"
Option Explicit


Public Function Encrypt(ByVal strUserCode As String, _
        strUserPassword) As String

Dim intCntr As Integer
Dim strEncyptedPassword As String
Dim strTruncatedchar As String
Dim lngAsciiNo As Double
    
    For intCntr = 1 To Len(strUserCode)
        lngAsciiNo = lngAsciiNo + Asc(Mid(strUserCode, intCntr, 1))
    Next
    
    Do While lngAsciiNo > 100
        If (lngAsciiNo Mod 2) Then
            lngAsciiNo = lngAsciiNo / 3
        Else
            lngAsciiNo = lngAsciiNo / 4
        End If
    Loop
    
    For intCntr = 1 To Len(strUserPassword)
            strTruncatedchar = Asc(Mid(strUserPassword, intCntr, 1)) + lngAsciiNo
            strEncyptedPassword = strEncyptedPassword & Chr(intCntr + strTruncatedchar) & Chr(strTruncatedchar)
    Next
    
    Encrypt = strEncyptedPassword
End Function

Public Function Decrypt(ByVal strUserCode As String, _
            ByVal strUserPassword As String) As String

Dim intCntr As Integer
Dim strEncyptedPassword As String
Dim strTruncatedchar As String
Dim lngAsciiNo As Double
Dim strTruncatedWord As String
Dim intCntrLoop As Integer
    
    For intCntr = 1 To Len(strUserCode)
        lngAsciiNo = lngAsciiNo + Asc(Mid(strUserCode, intCntr, 1))
    Next
    
    Do While lngAsciiNo > 100
        If (lngAsciiNo Mod 2) Then
            lngAsciiNo = lngAsciiNo / 3
        Else
            lngAsciiNo = lngAsciiNo / 4
        End If
    Loop
   
    For intCntr = 1 To (Len(strUserPassword) / 2)
            strTruncatedWord = strTruncatedWord & Mid(strUserPassword, intCntr * 2, 1)
    Next
    
    intCntrLoop = 0
    
    For intCntr = 1 To Len(strTruncatedWord)
            strTruncatedchar = Asc(Mid(strTruncatedWord, intCntr, 1)) - lngAsciiNo
            strEncyptedPassword = strEncyptedPassword & Chr(strTruncatedchar)
    Next
    
    Decrypt = strEncyptedPassword
End Function


