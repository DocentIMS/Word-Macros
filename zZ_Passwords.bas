Attribute VB_Name = "zZ_Passwords"
'Option Explicit
'
'Function DecryptPassword(ByVal pstrPasswordCode As String) As String
'    ' Original Code https://stackoverflow.com/questions/47990187/securely-store-password-in-a-vba-project?utm_medium=organic&utm_source=google_rich_qa&utm_campaign=google_rich_qa
'    ' Modified to extend password length
'    ' Modifications free to use
'    Dim codeLen As Integer
'
'    Dim intChar As Integer
'    Dim intCode As Integer
'    Dim arrintShifts() As Integer
'    Dim arrlngCharCode() As Long
'    Dim icp As Integer
'    If Len(pstrPasswordCode) = 0 Then Exit Function
'    On Error GoTo ex
'    ' Initialise Arrays
'    icp = IIf(Right(pstrPasswordCode, 1) Mod 2 = 0, 5, 4)
'    pstrPasswordCode = Left(pstrPasswordCode, Len(pstrPasswordCode) - IIf(Right(pstrPasswordCode, 1) Mod 2 = 0, 1, 1))
'    codeLen = Len(pstrPasswordCode) / icp - 1 ' Array Index starts with 0
'    ReDim arrintShifts(codeLen)
'    ReDim arrlngCharCode(codeLen)
'
'    intChar = 0
'    intCode = 0
'
'    For intCode = 0 To codeLen
'        'store -8 to -1 into 0-7
'        arrintShifts(intCode) = intCode - (codeLen + 1)
'    Next intCode
'
'    'the code is stored by using the number of the letter of the password in the 4th character.
'    'the real code of the character is directly behind that.
'    'so the code 30555112012321187051111661144119
'    'has on position 3, 055, 5, 112, 0, 123, 2, 118, 7, 051, 1, 116, 6, 114 and 4, 119
'    'so sorted this is 0, 123, 1, 116, 2, 118, 3, 055, 4, 119, 5, 112, 6, 114, 7, 051
'    'then there is also the part where those charcode are shifted by adding -8 to -1 to them.
'    'leading to the real charactercodes:
'    '0, 123-8, 1, 116-7, 2, 118-6, 3, 055-5, 4, 119-4, 5, 112-3, 6, 114-2, 7, 051-1
'    '0, 115, 1, 109, 2, 112, 3, 050, 4, 115, 5, 109, 6, 112, 7, 050
'
'    For intChar = 0 To codeLen
'        For intCode = 0 To codeLen
'            If CInt(Mid(pstrPasswordCode, intCode * icp + 1, icp - 3)) = intChar Then
'                arrlngCharCode(intChar) = (Mid(pstrPasswordCode, (intCode + 1) * icp - 2, 3) + arrintShifts(intChar))
'                Exit For
'            End If
'        Next intCode
'    Next intChar
'
'    'by getting the charcodes of these values, you create the password
'    DecryptPassword = ""
'    For intChar = 0 To codeLen
'        DecryptPassword = DecryptPassword & Chr(arrlngCharCode(intChar))
'    Next intChar
'ex:
'End Function
'
'Function EncryptPassword(ByVal pstrPasswordCode As String) As String
'    ' Generator free to use
'    Dim pwLen As Integer
'    Dim scp As String   ' String Code Position, for formatting "0" or "00"
'    Dim icp As Integer  ' marker if pwLen < 10 or > 10
'    Dim intCode As Integer
'    Dim arrintShifts() As Integer
'    Dim arrlngCharCode() As Long
'    Dim PW() As String
'
'    Dim Temp As Variant
'    Dim arnd() As Variant
'    Dim irnd As Variant
'    If Len(pstrPasswordCode) = 0 Then Exit Function
'    Randomize
'
'    ' Initialise Arrays
'    pwLen = Len(pstrPasswordCode) - 1 ' Array Index starts with 0
'    scp = IIf(pwLen < 10, "0", "00")
'    ' Create odd/even marker if we have 1 (odd) or 2 (even) byte index digits (scp), values between 0 and 9
'    icp = IIf(pwLen < 10, Int(Rnd() * 5 + 1) * 2 - 1, Int(Rnd() * 5 + 1) * 2)
'
'    ReDim arrintShifts(pwLen)
'    ReDim arrlngCharCode(pwLen)
'    ReDim PW(pwLen)
'    ReDim arnd(pwLen)
'
'    For intCode = 0 To pwLen
'        arnd(intCode) = intCode
'    Next intCode
'
'    ' randomize the indizes to bring the code into a random order
'    For intCode = LBound(arnd) To UBound(arnd)
'        irnd = CLng(((UBound(arnd) - intCode) * Rnd) + intCode)
'        If intCode <> irnd Then
'            Temp = arnd(intCode)
'            arnd(intCode) = arnd(irnd)
'            arnd(irnd) = Temp
'        End If
'    Next intCode
'
'    'by getting the charcodes of these values, you create the password
'    For intCode = 0 To pwLen
'        'get characters
'        PW(intCode) = Mid(pstrPasswordCode, intCode + 1, 1)
'        'and store -8 to -1 into 0-7 (for additional obfuscation)
'        arrintShifts(intCode) = intCode - (pwLen + 1)
'    Next intCode
'
'    ' Search for the random index and throw the shifted code at this position
'    For intCode = 0 To pwLen
'        arrlngCharCode(Match(intCode, arnd, False)) = AscB(PW(intCode)) - arrintShifts(intCode)
'    Next intCode
'
'    ' Chain All Codes, combination of arnd(intcode) and arrlngCharCode(intcode) gives the random order
'    EncryptPassword = ""
'    For intCode = 0 To pwLen
'        EncryptPassword = EncryptPassword & Format(arnd(intCode), scp) & Format(arrlngCharCode(intCode), "000")
'    Next intCode
'    EncryptPassword = EncryptPassword & icp
'
'End Function
