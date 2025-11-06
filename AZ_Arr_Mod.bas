Attribute VB_Name = "AZ_Arr_Mod"
Option Explicit
Option Compare Text
Option Private Module
#If VBA7 Then
  Private Type Pointer: value As LongPtr: End Type
  Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef src As Any, ByVal Size As LongPtr)
#Else
  Private Type Pointer: value As Long: End Type
  Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (ByRef dest As Any, ByRef src As Any, ByVal Size As Long)
#End If
Private Type TtagVARIANT
    vt As Integer
    r1 As Integer
    r2 As Integer
    r3 As Integer
    sa As Pointer
End Type
Function ArrToJSON(Arr As Variant, Optional TableName = "Table", Optional WithHeaders As Boolean) As String
    Dim s As String, r As Long, c As Long, Cs As String, ss() As String, i As Long
    Dim Rs As Long
    s = """" & TableName & """: ["
    On Error GoTo ex
    Rs = LBound(Arr)
    If WithHeaders Then Rs = Rs + 1
    For r = Rs To UBound(Arr)
        s = s & "{"
        For c = 1 To UBound(Arr, 2)
            Cs = Arr(r, c)
            If WithHeaders Then
                s = s & """" & Arr(Rs - 1, c) & """: "
            Else
                s = s & """Column " & c & """: "
            End If
            If InStr(Cs, Chr(13)) Then
                ss = Split(Cs, Chr(13))
                s = s & "["
                For i = 0 To UBound(ss)
                    s = s & """" & Trim(ss(i)) & ""","
                Next
                s = Left$(s, Len(s) - 1) & "]"
            Else
                s = s & """" & Cs & """"
            End If
            s = s & ","
        Next
        s = Left$(s, Len(s) - 1) & "},"
    Next
    ArrToJSON = Left$(s, Len(s) - IIf(Right$(s, 1) = ",", 1, 0)) & "]"
ex:
End Function
Public Function GetDims(Arr As Variant) As Integer
    Dim va As TtagVARIANT
    RtlMoveMemory va, Arr, LenB(va)                                               ' read tagVARIANT              '
    If va.vt And &H2000 Then Else Exit Function                                   ' exit if not an array         '
    If va.vt And &H4000 Then RtlMoveMemory va.sa, ByVal va.sa.value, LenB(va.sa)  ' read by reference            '
    If va.sa.value Then RtlMoveMemory GetDims, ByVal va.sa.value, 2               ' read cDims from tagSAFEARRAY '
End Function
Public Function GetDimsAndBounds(Arr As Variant) As Collection
    Dim Coll As New Collection, i As Long
    For i = 1 To GetDims(Arr)
        Coll.Add Array(LBound(Arr, i), UBound(Arr, i))
    Next
    Set GetDimsAndBounds = Coll
End Function
Function CollToDict(Coll As Collection, Optional Ky As String) As Dictionary
    Dim i As Long, Dict As New Dictionary
    On Error GoTo ex
    For i = 1 To Coll.Count
        If Len(Ky) Then
            Dict.Add Coll(i)(Ky), Coll(i)
        Else
            Dict.Add Coll(i), Coll(i)
        End If
    Next
ex:
    Set CollToDict = Dict
End Function
Function DictCollToArr(DictColl As Variant) As Variant
    Dim Arr As Variant, i As Long
    On Error GoTo ex
    ReDim Arr(1 To DictColl.Count)
    For i = 1 To DictColl.Count
        Arr(i) = DictColl.Item(i)
    Next
ex:
    DictCollToArr = Arr
End Function
Function ArrToDict(Arr As Variant, Optional AllowRepeting As Boolean) As Dictionary
    Dim Dict As New Dictionary, i As Long
    On Error GoTo ex
    If TypeName(Arr) = "Dictionary" Then
        Set Dict = Arr
    Else
        On Error Resume Next
        For i = LBound(Arr) To UBound(Arr)
            Dict.Add Arr(i), Arr(i)
            If Err.Number And AllowRepeting Then
                Dict.Add i & "_" & Arr(i), Arr(i)
                Err.Clear
            End If
        Next
    End If
ex:
    Set ArrToDict = Dict
End Function
Function ArrToColl(Arr As Variant, Optional AllowRepeting As Boolean) As Collection
    Dim Coll As New Collection, i As Long
    On Error GoTo ex
    If TypeName(Arr) = "Collection" Then
        Set Coll = Arr
    Else
        On Error Resume Next
        For i = LBound(Arr) To UBound(Arr)
            Coll.Add Arr(i), Arr(i)
            If Err.Number <> 0 And AllowRepeting Then
                Coll.Add Arr(i), i & "_" & Arr(i)
                Err.Clear
            End If
        Next
    End If
ex:
    Set ArrToColl = Coll
End Function
Function Filter2DArr(V, ByVal Arr, c, Optional Dm, Optional Rs)
    Dim i As Long, j As Long, tArr, x As Long, oArr()
    If IsMissing(Dm) Then Dm = 2
    If IsMissing(Rs) Then Rs = LBound(Arr, Dm)
    If Dm = 1 Then tArr = Transpose(Arr) Else tArr = Arr
    For i = LBound(tArr, 2) To UBound(tArr, 2)
        If tArr(CLng(c), i) Like V Or i < Rs Then
            x = x + 1
            ReDim Preserve oArr(LBound(tArr) To UBound(tArr), LBound(tArr, 2) To x)
            For j = LBound(tArr) To UBound(tArr)
                oArr(j, x) = tArr(j, i)
            Next
        End If
    Next
    If Dm = 1 Then Filter2DArr = Transpose(oArr) Else Filter2DArr = oArr
End Function
Function CollToArr(Coll As Collection) As Variant
    Dim Arr(), i As Long
    If Coll.Count Then ReDim Preserve Arr(1 To Coll.Count)
    On Error GoTo ex
    For i = 1 To Coll.Count
        If TypeName(Coll(i)) = "Dictionary" Then
            Arr(i) = Coll(i)("title")
        Else
            Arr(i) = Coll(i)
        End If
    Next
ex:
    CollToArr = Arr
End Function
Function CollTo2DArr(Coll As Collection) As Variant
    Dim Arr(), i As Long
    ReDim Preserve Arr(1 To Coll.Count, 1 To 1)
    On Error GoTo ex
    For i = 1 To Coll.Count
        Arr(i, 1) = Coll(i)
    Next
ex:
    CollTo2DArr = Arr
End Function
Function Transpose(Arr As Variant) As Variant
    Dim i As Long, j As Long
    On Error GoTo ex
    ReDim NArr(LBound(Arr, 2) To UBound(Arr, 2), LBound(Arr) To UBound(Arr))
    For i = LBound(Arr) To UBound(Arr)
        For j = LBound(Arr, 2) To UBound(Arr, 2)
            NArr(j, i) = Arr(i, j)
        Next
    Next
ex:
    Transpose = NArr
End Function
Function Match(What As Variant, Arr As Variant, Optional MatchMode As Long = 0) As Variant
    Dim i As Long, l As Variant, V As Variant ', x As Long
    On Error Resume Next
    If LBound(Arr) <> LBound(Arr) Then Exit Function
    Select Case MatchMode
    Case -1: V = 9999999999#
    Case 1: V = -9999999999#
    End Select
    For i = LBound(Arr) To UBound(Arr)
        Select Case MatchMode
        Case 0: If What = Arr(i) Then Match = i: Exit For
        Case -1: If What <= Arr(i) Then If V > Arr(i) Then V = Arr(i): Match = i 'Smallest Value Greater Than What
        Case 1: If What >= Arr(i) Then If V < Arr(i) Then V = Arr(i): Match = i 'Largest Value Less Than What
        End Select
    Next
End Function
Function Sort2DArrayMulti(Arr As Variant, Cs, Optional Rs As Long = -1, _
            Optional re As Long = -1, Optional Ascendings, Optional RemoveDuplicates As Boolean = False)
    Dim i As Long, c As Long, j As Long, oArr()

    If re = -1 Then re = UBound(Arr)
    If Rs = -1 Then Rs = LBound(Arr)
    i = LBound(Cs)
    If IsMissing(Ascendings) Then
        ReDim Ascendings(i To UBound(Cs)) As Boolean
        For i = i To UBound(Cs)
            Ascendings(i) = True
        Next
    End If
    i = LBound(Cs)
    Arr = Sort2DArray(Arr, Cs(i), Rs, re, Ascendings(i))
    For i = i + 1 To UBound(Cs)
        Arr = Sort2DArrayRespect(Arr, Cs, i, Rs, re, Ascendings(i))
    Next
    If RemoveDuplicates Then
        
        j = LBound(Arr)
        For i = j + 1 To UBound(Arr) 'For i = LBound(Arr) To UBound(Arr) - 1
            If Arr(i - 1, Cs(LBound(Cs))) <> Arr(i, Cs(LBound(Cs))) Then j = j + 1
        Next
        ReDim Preserve oArr(LBound(Arr) To j, LBound(Arr, 2) To UBound(Arr, 2))
        j = LBound(Arr)
        For c = LBound(Arr, 2) To UBound(Arr, 2)
            oArr(j, c) = Arr(j, c)
        Next
        For i = j + 1 To UBound(Arr)
            If Arr(i - 1, Cs(LBound(Cs))) <> Arr(i, Cs(LBound(Cs))) Then
                j = j + 1
                For c = LBound(Arr, 2) To UBound(Arr, 2)
                    oArr(j, c) = Arr(i, c)
                Next
            End If
        Next i
        Sort2DArrayMulti = oArr
    Else
        Sort2DArrayMulti = Arr
    End If
End Function
Private Function Sort2DArrayRespect(Arr As Variant, Cs, Optional cRespect = -1, Optional Rs As Long = -1, Optional re As Long = -1, Optional Ascending = True)
    Dim i As Long, j As Long, x As Long, Temp, DoSwitch As Boolean, c As Long, ci As Long
    If re = -1 Then re = UBound(Arr)
    If Rs = -1 Then Rs = LBound(Arr)
    If cRespect = -1 Then cRespect = LBound(Cs)
    c = Cs(cRespect)
    x = Rs
    For i = Rs To re - 1
        DoSwitch = False
        For ci = LBound(Cs) To cRespect - 1
            DoSwitch = DoSwitch Or Arr(i, CInt(Cs(ci))) <> Arr(i + 1, CInt(Cs(ci)))
            If DoSwitch Then Exit For
        Next
        If DoSwitch Then
            Arr = Sort2DArray(Arr, c, x, i, Ascending)
            x = i + 1
        End If
    Next
    On Error GoTo ex
    If Arr(i - 1, CInt(cRespect)) = Arr(i, CInt(cRespect)) Then
        Arr = Sort2DArray(Arr, c, x, i, Ascending)
    End If
ex:
    Sort2DArrayRespect = Arr
End Function
Function Sort2DArray(Arr As Variant, c, Optional Rs As Long = -1, Optional re As Long = -1, Optional Ascending = True)
    Dim i As Long, j As Long, x As Long, Temp, DoSwitch As Boolean
    On Error GoTo ex
    If re = -1 Then re = UBound(Arr)
    If Rs = -1 Then Rs = LBound(Arr)
    For i = Rs To re - 1
        For j = i + 1 To re
            If Ascending Then
                DoSwitch = Arr(i, CInt(c)) > Arr(j, CInt(c))
            Else
                DoSwitch = Arr(i, CInt(c)) < Arr(j, CInt(c))
            End If
            If DoSwitch Then
                For x = LBound(Arr, 2) To UBound(Arr, 2)
                    Temp = Arr(j, x)
                    Arr(j, x) = Arr(i, x)
                    Arr(i, x) = Temp
                Next
            End If
        Next j
    Next i
ex:
    Sort2DArray = Arr
End Function

Function MaxIn2DArray(ByVal Arr As Variant, Optional cNum As Long = 1, Optional DimNum As Long = 1, Optional Min As Boolean = False, _
    Optional LBnd As Long, Optional UBnd As Long, Optional After As Variant, Optional ReturnPos As Boolean = False) As Variant
    Dim i As Long
    Dim A As Double
    Dim p As Long

    If LBnd = 0 Then LBnd = LBound(Arr, DimNum)
    If UBnd = 0 Then UBnd = UBound(Arr, DimNum)
    If UBnd < LBnd Then: i = LBnd: LBnd = UBnd: UBnd = i
    On Error Resume Next
    If Min Then A = 1 / 0 Else A = -1 / 0
    MaxIn2DArray = A
    If Not IsMissing(After) Then A = After
    Select Case DimNum
        Case 1
            If Min Then
                For i = LBnd To UBnd
                    If Arr(i, cNum) < MaxIn2DArray And Arr(i, cNum) > A Then
                        MaxIn2DArray = Arr(i, cNum)
                        p = i
                    End If
                Next
            Else
                For i = LBnd To UBnd
                    If Arr(i, cNum) > MaxIn2DArray And Arr(i, cNum) < A Then
                        MaxIn2DArray = Arr(i, cNum)
                        p = i
                    End If
                Next
            End If
            If ReturnPos Then MaxIn2DArray = p
        Case 2
            If Min Then
                For i = LBnd To UBnd
                    If Arr(cNum, i) > MaxIn2DArray And Arr(cNum, i) < A Then
                        MaxIn2DArray = Arr(cNum, i)
                        p = i
                    End If
                Next
            Else
                For i = LBnd To UBnd
                    If Arr(cNum, i) > MaxIn2DArray And Arr(cNum, i) > A Then
                        MaxIn2DArray = Arr(cNum, i)
                        p = i
                    End If
                Next
            End If
            If ReturnPos Then MaxIn2DArray = p
    End Select
End Function
Private Function SortByScores(Arr As Variant) As Variant
    SortByScores = Sort2DArrayMulti(Arr, Array(2, 3, 4), , , Array(False, False, True))
End Function
Function Sort1DArray(Arr As Variant, Optional ByVal Ascending As Boolean = True, Optional RemoveDuplicates As Boolean = False) As Variant
    Dim i As Long, j As Long, Temp As Variant, DoSwitch As Boolean, oArr()
    For i = LBound(Arr) To UBound(Arr) - 1
        For j = i + 1 To UBound(Arr)
            If Ascending Then
                DoSwitch = Arr(i) > Arr(j)
            Else
                DoSwitch = Arr(i) < Arr(j)
            End If
            If DoSwitch Then
                Temp = Arr(j)
                Arr(j) = Arr(i)
                Arr(i) = Temp
            End If
        Next j
    Next i
    If RemoveDuplicates Then
        j = 1
        ReDim Preserve oArr(1 To j)
        oArr(1) = Arr(LBound(Arr))
        For i = LBound(Arr) + 1 To UBound(Arr)
            If Arr(i - 1) <> Arr(i) Then
                j = j + 1
                ReDim Preserve oArr(1 To j)
                oArr(j) = Arr(i)
            End If
        Next i
        Sort1DArray = oArr
    Else
        Sort1DArray = Arr
    End If
End Function

'Function Sort2DArrayMulti(Arr As Variant, cs As Variant, Optional rs As Long = -1, _
'    Optional Re As Long = -1, Optional Ascendings As Variant)
'    Dim i As Long
'
'
'    If Re = -1 Then Re = UBound(Arr)
'    If rs = -1 Then rs = LBound(Arr)
'    i = LBound(cs)
'    If IsMissing(Ascendings) Then
'        ReDim Ascendings(i To UBound(cs)) As Boolean
'        For i = i To UBound(cs)
'            Ascendings(i) = True
'        Next
'    End If
'    i = LBound(cs)
'    Arr = Sort2DArray(Arr, cs(i), rs, Re, Ascendings(i))
'    For i = i + 1 To UBound(cs)
'        Arr = Sort2DArrayRespect(Arr, cs, i, rs, Re, Ascendings(i))
'        '        Arr = Sort2DArrayRespect(Arr, cs(i), cs(i - 1), rs, Re, Ascendings(i))
'    Next
'    Sort2DArrayMulti = Arr
'End Function
'Private Function Sort2DArrayRespect(Arr As Variant, cs As Variant, Optional cRespect As Variant = -1, Optional rs As Long = -1, Optional Re As Long = -1, Optional Ascending As Variant = True)
'    Dim i As Long
'    Dim x As Long
'    Dim DoSwitch As Boolean
'    Dim c As Long
'    Dim ci As Long
'
'    If Re = -1 Then Re = UBound(Arr)
'    If rs = -1 Then rs = LBound(Arr)
'    If cRespect = -1 Then cRespect = LBound(cs)
'    c = cs(cRespect)
'    x = rs
'    For i = rs To Re - 1
'        DoSwitch = False
'        For ci = LBound(cs) To cRespect - 1
'            DoSwitch = DoSwitch Or Arr(i, CInt(cs(ci))) <> Arr(i + 1, CInt(cs(ci)))
'            If DoSwitch Then Exit For
'        Next
'        '        If Arr(i, CInt(cRespect)) <> Arr(i + 1, CInt(cRespect)) Then
'        If DoSwitch Then
'            Arr = Sort2DArray(Arr, c, x, i, Ascending)
'            x = i + 1
'        End If
'    Next
'    If Arr(i - 1, CInt(cRespect)) = Arr(i, CInt(cRespect)) Then
'        Arr = Sort2DArray(Arr, c, x, i, Ascending)
'    End If
'    Sort2DArrayRespect = Arr
'End Function
'Function Sort2DArray(Arr As Variant, c As Variant, Optional rs As Long = -1, Optional Re As Long = -1, Optional Ascending As Variant = True)
'    Dim i As Long
'    Dim j As Long
'    Dim x As Long
'    Dim Temp As Variant
'    Dim DoSwitch As Boolean
'
'    If Re = -1 Then Re = UBound(Arr)
'    If rs = -1 Then rs = LBound(Arr)
'    For i = rs To Re - 1
'        For j = i + 1 To Re
'            If Ascending Then
'                DoSwitch = Arr(i, CInt(c)) > Arr(j, CInt(c))
'            Else
'                DoSwitch = Arr(i, CInt(c)) < Arr(j, CInt(c))
'            End If
'            If DoSwitch Then
'                For x = LBound(Arr, 2) To UBound(Arr, 2)
'                    Temp = Arr(j, x)
'                    Arr(j, x) = Arr(i, x)
'                    Arr(i, x) = Temp
'                Next
'            End If
'        Next j
'    Next i
'    Sort2DArray = Arr
'End Function
'@EntryPoint
'Function Sort1DArray(Arr As Variant, Optional Ascending As Boolean = True)
'    Dim i As Long
'    Dim j As Long
'    Dim Temp As Variant
'    Dim DoSwitch As Boolean
'
'    For i = LBound(Arr) To UBound(Arr) - 1
'        For j = i + 1 To UBound(Arr)
'            If Ascending Then
'                DoSwitch = Arr(i) > Arr(j)
'            Else
'                DoSwitch = Arr(i) < Arr(j)
'            End If
'            If DoSwitch Then
'                Temp = Arr(j)
'                Arr(j) = Arr(i)
'                Arr(i) = Temp
'            End If
'        Next j
'    Next i
'    Sort1DArray = Arr
'End Function



