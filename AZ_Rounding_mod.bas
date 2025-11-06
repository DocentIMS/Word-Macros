Attribute VB_Name = "AZ_Rounding_mod"
Option Explicit
Private Const xlErrValue = 2015
'2025-01-22 04:03 AM
Function RndDown(Amount As Variant, Optional Digits As Integer = 0) As Variant
    If Not IsNumeric(Amount) Then RndDown = CVErr(xlErrValue): Exit Function
    RndDown = Int(Amount * (10 ^ Digits)) / (10 ^ Digits)
End Function
Function RndUp(Amount As Variant, Optional Digits As Integer = 0) As Variant
    If Not IsNumeric(Amount) Then RndUp = CVErr(xlErrValue): Exit Function
    RndUp = Round((Amount + (5 / (10 ^ (Digits + 1)))) * (10 ^ Digits)) / (10 ^ Digits)
End Function
'@EntryPoint
Function RoundUpx(n As Single, Optional x As Long = 10) As Single
    RoundUpx = RndUp(n / x) * x
End Function
'@EntryPoint
Function RoundDownx(n As Single, Optional x As Long = 10) As Single
    RoundDownx = RndDown(n / x) * x
End Function
'@EntryPoint
Function Rndx(Amount As Double, Optional Digits As Integer = 0, Optional MidNum As Integer = 5) As Variant
    If Not IsNumeric(Amount) Then Rndx = CVErr(xlErrValue): Exit Function
    Dim AStr As String
    Dim n As Long
    Dim i As Long
    Dim s As String

    AStr = Amount
    i = InStr(AStr, "E")
    If i Then
        n = Right$(AStr, Len(AStr) - i)
        AStr = Left$(AStr, i - 1)
        i = InStr(AStr, ".")
        If i > 0 Then i = i - 1
        i = i + n
        AStr = Replace(AStr, ".", vbNullString)
        If i > 0 Then
            AStr = AStr & Zeros(i)
        ElseIf i < 0 Then
            AStr = Zeros(-i) & AStr
            AStr = "0." & AStr
        Else
            AStr = "0." & AStr
        End If
    End If
    If MidNum <> 5 Then
        i = InStr(AStr, ".") + Digits
        If Digits < 0 Then
            If i > 0 Then
                Rndx = Val(Left$(AStr, i - 1))
                s = Mid$(AStr, i + 1, Len(MidNum))
                If InStr(s, ".") Then s = Mid$(AStr, i, Len(MidNum))
            End If
            Rndx = Rndx & Zeros(-Digits)
        ElseIf i > 0 Then
            Rndx = Val(Left$(AStr, i))
            If i + Len(MidNum) >= Len(AStr) Then AStr = AStr & Zeros(i + Len(MidNum) - Len(AStr))
            s = Mid$(AStr, i + 1, Len(MidNum))
            If InStr(s, ".") Then s = Mid$(AStr, i + 2, Len(MidNum))
        Else
            Rndx = AStr
        End If
        If s <> vbNullString Then
            If s >= MidNum Then
                Rndx = Rndx + (10 ^ (-Digits))
            End If
        End If
    Else
        Rndx = Round((AStr) * (10 ^ Digits)) / (10 ^ Digits)
    End If
End Function
Function Zeros(n As Long, Optional Str As String = "0") As String
    Zeros = Replace(Space$(n), " ", Str)
End Function
