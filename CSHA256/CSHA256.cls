VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSHA256"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ModuleWithoutFolder, UseMeaningfulName, AssignedByValParameter

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
'
'                                      CSHA256
'
' SHA-2 256 hashing algorithm class.
'
' For documentation, licensing, updates etc. see:
' <https://github.com/kimbar/VBA-Native-Tools>
'
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

Option Explicit

Private hash(0 To 7) As Long
Private w(0 To 63) As Long
Private total_bits_consumed_ll As Long
Private total_bits_consumed_hl As Long
Private current_chunk_bytes_taken As Byte
Private IsFinished As Boolean

' Consts
Private two_topmost_bits_values(0 To 3) As Long
Private three_topmost_bits_values(0 To 7) As Long
Private sha_k(0 To 63) As Long

Private Sub Class_Initialize()
    ' constants needed to simulate unsigned arithmetic
    PopulateLongArrayFromString two_topmost_bits_values, _
        "00000000 40000000 80000000 C0000000"
    PopulateLongArrayFromString three_topmost_bits_values, _
        "00000000 20000000 40000000 60000000 80000000 A0000000 C0000000 E0000000"
    ' SHA-2 256 constants
    PopulateLongArrayFromString sha_k, _
        "428A2F98 71374491 B5C0FBCF E9B5DBA5 3956C25B 59F111F1 923F82A4 AB1C5ED5 " & _
        "D807AA98 12835B01 243185BE 550C7DC3 72BE5D74 80DEB1FE 9BDC06A7 C19BF174 " & _
        "E49B69C1 EFBE4786 0FC19DC6 240CA1CC 2DE92C6F 4A7484AA 5CB0A9DC 76F988DA " & _
        "983E5152 A831C66D B00327C8 BF597FC7 C6E00BF3 D5A79147 06CA6351 14292967 " & _
        "27B70A85 2E1B2138 4D2C6DFC 53380D13 650A7354 766A0ABB 81C2C92E 92722C85 " & _
        "A2BFE8A1 A81A664B C24B8B70 C76C51A3 D192E819 D6990624 F40E3585 106AA070 " & _
        "19A4C116 1E376C08 2748774C 34B0BCB5 391C0CB3 4ED8AA4A 5B9CCA4F 682E6FF3 " & _
        "748F82EE 78A5636F 84C87814 8CC70208 90BEFFFA A4506CEB BEF9A3F7 C67178F2 "
    ' Reseting variables (state)
    Me.Reset
End Sub

Private Sub Class_Terminate()
    ' We generaly try to clear the internal state when not needed anymore to avoid spewing the data in memory
    Me.Reset
End Sub

Public Sub Reset()
    Dim idx As Long
    For idx = 0 To 63
        w(idx) = 0&
    Next
    ' Initial hash as stated in the specification
    PopulateLongArrayFromString hash, _
        "6A09E667 BB67AE85 3C6EF372 A54FF53A 510E527F 9B05688C 1F83D9AB 5BE0CD19"
    ' Length of the data (in bits) kept in two Longs (more than 2GB of data is possible this way)
    total_bits_consumed_ll = 0&
    total_bits_consumed_hl = 0&
    ' Number of bytes currently written to the buffer (chunk)
    current_chunk_bytes_taken = 0
    IsFinished = False
End Sub

Public Sub UpdateLong(ByVal data As Long)
    If (current_chunk_bytes_taken And 3) = 0 Then
        UpdateLongAligned data
    Else
        Err.Raise 1000, "Not implemented"
    End If
End Sub

Public Sub UpdateByte(ByVal data As Byte)
    ' The byte is "ored" to the buffer (chunk) after being l-shifted into position in Long
    w(current_chunk_bytes_taken \ 4) = w(current_chunk_bytes_taken \ 4) Or ByteToLong(data, current_chunk_bytes_taken And 3)
    current_chunk_bytes_taken = current_chunk_bytes_taken + 1
    IncrementTotalLength bits:=8
    ' Churn the chunk if full (16 Longs, 64 bytes)
    If current_chunk_bytes_taken = 64 Then
        ChurnTheChunk
        current_chunk_bytes_taken = 0
    End If
End Sub

Private Sub UpdateBytesArrayNaive(ByRef data() As Byte, ByVal lb As Long, ByVal ub As Long)
    ' Naive implementation, only to be used for unaligned ends of arrays
    Dim idx As Long
    For idx = lb To ub
        UpdateByte data(idx)
    Next
End Sub

Public Sub UpdateBytesArray(ByRef data() As Byte)
    Dim start_run As Byte
    Dim cursor As Long: cursor = LBound(data)
    Dim ub As Long: ub = UBound(data)
    Dim length As Long: length = ub - cursor + 1

    ' Current state of the buffer (chunk) may be unaligned to the Long, here we determine if we need to first realign
    ' the buffer. An edge case - data is to short to realign the buffer - is also taken into account
    start_run = current_chunk_bytes_taken Mod 4
    If start_run Then start_run = 4 - start_run
    If length < start_run Then start_run = length
    UpdateBytesArrayNaive data, cursor, cursor + start_run - 1
    cursor = cursor + start_run
    ' Now we are long-aligned (unless it was imposible, because data was too short)
    If cursor > ub Then Exit Sub
    ' Efficient implementation since we're aligned to the Long
    Do While ub - cursor >= 4
        UpdateLongAligned FourBytesToLong(data(cursor), data(cursor + 1), data(cursor + 2), data(cursor + 3))
        cursor = cursor + 4
    Loop
    ' If there is some leftover, we put it in the buffer the naive way
    UpdateBytesArrayNaive data, cursor, ub
End Sub

Private Sub UpdateLongAligned(ByVal data As Long)
    ' Fast implementation - only assignment of a single variable to the buffer (chunk)
    w(current_chunk_bytes_taken \ 4) = data
    current_chunk_bytes_taken = current_chunk_bytes_taken + 4
    IncrementTotalLength bits:=32
    ' Churn the chunk if full (16 Longs, 64 bytes)
    If current_chunk_bytes_taken = 64 Then
        ChurnTheChunk
        current_chunk_bytes_taken = 0
    End If
End Sub

Private Sub IncrementTotalLength(ByVal bits As Long, Optional ByVal len_4Gbits As Long = 0&)
    Dim c As Boolean
    total_bits_consumed_ll = UnsigSumWithCarry(total_bits_consumed_ll, bits, c, c)
    ' This line is only ever used with data larger than 2GB
    If c Or len_4Gbits Then total_bits_consumed_hl = UnsigSumWithCarry(total_bits_consumed_hl, len_4Gbits, c, c)
End Sub

Public Sub Finish()
    ' Setting up the `1` bit just behind the data, as specification states
    w(current_chunk_bytes_taken \ 4) = w(current_chunk_bytes_taken \ 4) Or ByteToLong(&H80, current_chunk_bytes_taken Mod 4)
    ' The specification claims 65 bits for padding. Since we're operating at a byte level that is 9 bytes
    ' If in curren chunk we used 55 or less bytes then we can fit the padding in the current chunk
    ' If not we need to churn the chunk (with the `1` bit behind the data - that's required always) and produce
    ' anoyher chunk just for the length field in padding
    If current_chunk_bytes_taken > 55 Then ChurnTheChunk
    w(14) = total_bits_consumed_hl
    w(15) = total_bits_consumed_ll
    ' In any case we need to churn the last chunk
    ChurnTheChunk
    Dim idx As Long
    ' Clearing the chunk just to be extra safe
    For idx = 16 To 63
        w(idx) = 0&
    Next
    IsFinished = True
End Sub

Friend Sub ChurnTheChunk()

    If IsFinished Then GoTo skip

    Dim idx As Long
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim e As Long
    Dim f As Long
    Dim g As Long
    Dim h As Long
    Dim t1_ As Long
    Dim t2_ As Long

    a = hash(0)
    b = hash(1)
    c = hash(2)
    d = hash(3)
    e = hash(4)
    f = hash(5)
    g = hash(6)
    h = hash(7)
    ' Here be dragons
    For idx = 0 To 63
        If idx > 15 Then w(idx) = UnsigSum4(w(idx - 16), Sigma0Expand(w(idx - 15)), w(idx - 7), Sigma1Expand(w(idx - 2)))
        t1_ = UnsigSum(UnsigSum4(h, Sigma1Normal(e), (e And f) Xor ((Not e) And g), sha_k(idx)), w(idx))
        t2_ = UnsigSum(Sigma0Normal(a), (a And b) Xor (a And c) Xor (b And c))
        h = g
        g = f
        f = e
        e = UnsigSum(d, t1_)
        d = c
        c = b
        b = a
        a = UnsigSum(t1_, t2_)
    Next

    hash(0) = UnsigSum(hash(0), a)
    hash(1) = UnsigSum(hash(1), b)
    hash(2) = UnsigSum(hash(2), c)
    hash(3) = UnsigSum(hash(3), d)
    hash(4) = UnsigSum(hash(4), e)
    hash(5) = UnsigSum(hash(5), f)
    hash(6) = UnsigSum(hash(6), g)
    hash(7) = UnsigSum(hash(7), h)
    ' We clear the buffer (chunk) (at least the original data part) right away beacause this simplifies reasoning
    ' in al of the `Update*` subs.
skip:
    For idx = 0 To 15
        w(idx) = 0&
    Next
End Sub

Public Function Digest() As String
    If Not IsFinished Then Me.Finish
    Dim i As Long
    Dim h As String
    For i = 0 To 7
        h = Hex$(hash(i))
        Digest = Digest & Left$("00000000", 8 - Len(h)) & h
    Next
End Function

' Unsigned arithmetic simulation
' ------------------------------

Friend Function UnsigSum(ByVal x As Long, ByVal y As Long) As Long
    Dim topmost_bits As Long
    If x And &H80000000 Then topmost_bits = topmost_bits + 2&
    If x And &H40000000 Then topmost_bits = topmost_bits + 1&
    If y And &H80000000 Then topmost_bits = topmost_bits + 2&
    If y And &H40000000 Then topmost_bits = topmost_bits + 1&
    UnsigSum = (x And &H3FFFFFFF) + (y And &H3FFFFFFF)
    If UnsigSum And &H40000000 Then
        topmost_bits = topmost_bits + 1&
        UnsigSum = UnsigSum And &H3FFFFFFF
    End If
    UnsigSum = UnsigSum Or two_topmost_bits_values(topmost_bits And 3&)
End Function

Friend Function UnsigSum4(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long) As Long
    Dim topmost_bits As Long
    If a And &H80000000 Then topmost_bits = topmost_bits + 4&
    If a And &H40000000 Then topmost_bits = topmost_bits + 2&
    If a And &H20000000 Then topmost_bits = topmost_bits + 1&
    If b And &H80000000 Then topmost_bits = topmost_bits + 4&
    If b And &H40000000 Then topmost_bits = topmost_bits + 2&
    If b And &H20000000 Then topmost_bits = topmost_bits + 1&
    If c And &H80000000 Then topmost_bits = topmost_bits + 4&
    If c And &H40000000 Then topmost_bits = topmost_bits + 2&
    If c And &H20000000 Then topmost_bits = topmost_bits + 1&
    If d And &H80000000 Then topmost_bits = topmost_bits + 4&
    If d And &H40000000 Then topmost_bits = topmost_bits + 2&
    If d And &H20000000 Then topmost_bits = topmost_bits + 1&
    UnsigSum4 = (a And &H1FFFFFFF) + (b And &H1FFFFFFF) + (c And &H1FFFFFFF) + (d And &H1FFFFFFF)
    If UnsigSum4 And &H40000000 Then topmost_bits = topmost_bits + 2&
    If UnsigSum4 And &H20000000 Then topmost_bits = topmost_bits + 1&
    UnsigSum4 = UnsigSum4 And &H1FFFFFFF
    UnsigSum4 = UnsigSum4 Or three_topmost_bits_values(topmost_bits And 7&)
End Function

Private Function UnsigSumWithCarry(ByVal x As Long, ByVal y As Long, ByVal z As Boolean, ByRef c As Boolean) As Long
    Dim topmost_bits As Long
    If x And &H80000000 Then topmost_bits = topmost_bits + 2&
    If x And &H40000000 Then topmost_bits = topmost_bits + 1&
    If y And &H80000000 Then topmost_bits = topmost_bits + 2&
    If y And &H40000000 Then topmost_bits = topmost_bits + 1&
    x = x And &H3FFFFFFF
    y = y And &H3FFFFFFF
    UnsigSumWithCarry = x + y - z    ' `z` is cast to `-1` if `True` or `0` if `False`
    If UnsigSumWithCarry And &H40000000 Then
        topmost_bits = topmost_bits + 1&
        UnsigSumWithCarry = UnsigSumWithCarry And &H3FFFFFFF
    End If
    UnsigSumWithCarry = UnsigSumWithCarry Or two_topmost_bits_values(topmost_bits And 3&)
    ' Max value of `topmost_bits` is (binary) 111
    c = topmost_bits \ 4&
End Function

Private Function ByteToLong(ByVal x As Byte, ByVal byte_idx_bigendian As Byte) As Long
    Select Case byte_idx_bigendian
        Case 3: ByteToLong = x
        Case 2: ByteToLong = x * &H100&
        Case 1: ByteToLong = x * &H10000
        Case 0:
            If (x And &H80) Then
                ByteToLong = ((x And &H7F) * &H1000000) Or &H80000000
            Else
                ByteToLong = x * &H1000000
            End If
    End Select
End Function

Private Function FourBytesToLong(ByVal hhbyte As Byte, ByVal lhbyte As Byte, ByVal hlbyte As Byte, ByVal llbyte As Byte) As Long
    Dim lower As Long
    lower = (lhbyte * &H10000) Or (hlbyte * &H100&) Or llbyte
    If (hhbyte And &H80) Then
        FourBytesToLong = ((hhbyte And &H7F) * &H1000000) Or &H80000000 Or lower
    Else
        FourBytesToLong = hhbyte * &H1000000 Or lower
    End If
End Function

Friend Function Sigma0Expand(ByVal x As Long) As Long
    Dim r As Long
    If x And &H1& Then r = r Xor &H2004000
    If x And &H2& Then r = r Xor &H4008000
    If x And &H4& Then r = r Xor &H8010000
    If x And &H8& Then r = r Xor &H10020001
    If x And &H10& Then r = r Xor &H20040002
    If x And &H20& Then r = r Xor &H40080004
    If x And &H40& Then r = r Xor &H80100008
    If x And &H80& Then r = r Xor &H200011
    If x And &H100& Then r = r Xor &H400022
    If x And &H200& Then r = r Xor &H800044
    If x And &H400& Then r = r Xor &H1000088
    If x And &H800& Then r = r Xor &H2000110
    If x And &H1000& Then r = r Xor &H4000220
    If x And &H2000& Then r = r Xor &H8000440
    If x And &H4000& Then r = r Xor &H10000880
    If x And &H8000& Then r = r Xor &H20001100
    If x And &H10000 Then r = r Xor &H40002200
    If x And &H20000 Then r = r Xor &H80004400
    If x And &H40000 Then r = r Xor &H8801&
    If x And &H80000 Then r = r Xor &H11002
    If x And &H100000 Then r = r Xor &H22004
    If x And &H200000 Then r = r Xor &H44008
    If x And &H400000 Then r = r Xor &H88010
    If x And &H800000 Then r = r Xor &H110020
    If x And &H1000000 Then r = r Xor &H220040
    If x And &H2000000 Then r = r Xor &H440080
    If x And &H4000000 Then r = r Xor &H880100
    If x And &H8000000 Then r = r Xor &H1100200
    If x And &H10000000 Then r = r Xor &H2200400
    If x And &H20000000 Then r = r Xor &H4400800
    If x And &H40000000 Then r = r Xor &H8801000
    If x And &H80000000 Then r = r Xor &H11002000
    Sigma0Expand = r
End Function

Friend Function Sigma1Expand(ByVal x As Long) As Long
    Dim r As Long
    If x And &H1& Then r = r Xor &HA000&
    If x And &H2& Then r = r Xor &H14000
    If x And &H4& Then r = r Xor &H28000
    If x And &H8& Then r = r Xor &H50000
    If x And &H10& Then r = r Xor &HA0000
    If x And &H20& Then r = r Xor &H140000
    If x And &H40& Then r = r Xor &H280000
    If x And &H80& Then r = r Xor &H500000
    If x And &H100& Then r = r Xor &HA00000
    If x And &H200& Then r = r Xor &H1400000
    If x And &H400& Then r = r Xor &H2800001
    If x And &H800& Then r = r Xor &H5000002
    If x And &H1000& Then r = r Xor &HA000004
    If x And &H2000& Then r = r Xor &H14000008
    If x And &H4000& Then r = r Xor &H28000010
    If x And &H8000& Then r = r Xor &H50000020
    If x And &H10000 Then r = r Xor &HA0000040
    If x And &H20000 Then r = r Xor &H40000081
    If x And &H40000 Then r = r Xor &H80000102
    If x And &H80000 Then r = r Xor &H205&
    If x And &H100000 Then r = r Xor &H40A&
    If x And &H200000 Then r = r Xor &H814&
    If x And &H400000 Then r = r Xor &H1028&
    If x And &H800000 Then r = r Xor &H2050&
    If x And &H1000000 Then r = r Xor &H40A0&
    If x And &H2000000 Then r = r Xor &H8140&
    If x And &H4000000 Then r = r Xor &H10280
    If x And &H8000000 Then r = r Xor &H20500
    If x And &H10000000 Then r = r Xor &H40A00
    If x And &H20000000 Then r = r Xor &H81400
    If x And &H40000000 Then r = r Xor &H102800
    If x And &H80000000 Then r = r Xor &H205000
    Sigma1Expand = r
End Function

Friend Function Sigma0Normal(ByVal x As Long) As Long
    Dim r As Long
    If x And &H1& Then r = r Xor &H40080400
    If x And &H2& Then r = r Xor &H80100800
    If x And &H4& Then r = r Xor &H201001
    If x And &H8& Then r = r Xor &H402002
    If x And &H10& Then r = r Xor &H804004
    If x And &H20& Then r = r Xor &H1008008
    If x And &H40& Then r = r Xor &H2010010
    If x And &H80& Then r = r Xor &H4020020
    If x And &H100& Then r = r Xor &H8040040
    If x And &H200& Then r = r Xor &H10080080
    If x And &H400& Then r = r Xor &H20100100
    If x And &H800& Then r = r Xor &H40200200
    If x And &H1000& Then r = r Xor &H80400400
    If x And &H2000& Then r = r Xor &H800801
    If x And &H4000& Then r = r Xor &H1001002
    If x And &H8000& Then r = r Xor &H2002004
    If x And &H10000 Then r = r Xor &H4004008
    If x And &H20000 Then r = r Xor &H8008010
    If x And &H40000 Then r = r Xor &H10010020
    If x And &H80000 Then r = r Xor &H20020040
    If x And &H100000 Then r = r Xor &H40040080
    If x And &H200000 Then r = r Xor &H80080100
    If x And &H400000 Then r = r Xor &H100201
    If x And &H800000 Then r = r Xor &H200402
    If x And &H1000000 Then r = r Xor &H400804
    If x And &H2000000 Then r = r Xor &H801008
    If x And &H4000000 Then r = r Xor &H1002010
    If x And &H8000000 Then r = r Xor &H2004020
    If x And &H10000000 Then r = r Xor &H4008040
    If x And &H20000000 Then r = r Xor &H8010080
    If x And &H40000000 Then r = r Xor &H10020100
    If x And &H80000000 Then r = r Xor &H20040200
    Sigma0Normal = r
End Function

Friend Function Sigma1Normal(ByVal x As Long) As Long
    Dim r As Long
    If x And &H1& Then r = r Xor &H4200080
    If x And &H2& Then r = r Xor &H8400100
    If x And &H4& Then r = r Xor &H10800200
    If x And &H8& Then r = r Xor &H21000400
    If x And &H10& Then r = r Xor &H42000800
    If x And &H20& Then r = r Xor &H84001000
    If x And &H40& Then r = r Xor &H8002001
    If x And &H80& Then r = r Xor &H10004002
    If x And &H100& Then r = r Xor &H20008004
    If x And &H200& Then r = r Xor &H40010008
    If x And &H400& Then r = r Xor &H80020010
    If x And &H800& Then r = r Xor &H40021
    If x And &H1000& Then r = r Xor &H80042
    If x And &H2000& Then r = r Xor &H100084
    If x And &H4000& Then r = r Xor &H200108
    If x And &H8000& Then r = r Xor &H400210
    If x And &H10000 Then r = r Xor &H800420
    If x And &H20000 Then r = r Xor &H1000840
    If x And &H40000 Then r = r Xor &H2001080
    If x And &H80000 Then r = r Xor &H4002100
    If x And &H100000 Then r = r Xor &H8004200
    If x And &H200000 Then r = r Xor &H10008400
    If x And &H400000 Then r = r Xor &H20010800
    If x And &H800000 Then r = r Xor &H40021000
    If x And &H1000000 Then r = r Xor &H80042000
    If x And &H2000000 Then r = r Xor &H84001
    If x And &H4000000 Then r = r Xor &H108002
    If x And &H8000000 Then r = r Xor &H210004
    If x And &H10000000 Then r = r Xor &H420008
    If x And &H20000000 Then r = r Xor &H840010
    If x And &H40000000 Then r = r Xor &H1080020
    If x And &H80000000 Then r = r Xor &H2100040
    Sigma1Normal = r
End Function

' Utils
' ------------------------------

Private Sub PopulateLongArrayFromString(ByRef arr() As Long, ByRef data As String)
    Dim idx As Long
    Dim cursor As Long: cursor = 1
    For idx = LBound(arr) To UBound(arr)
        arr(idx) = CLng("&H" & Mid$(data, cursor, 8))
        cursor = cursor + 9
    Next
End Sub

