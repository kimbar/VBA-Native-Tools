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

Private hash As Variant
Private w(0 To 63) As Long
Private total_bits_consumed_ll As Long
Private total_bits_consumed_hl As Long
Private current_chunk_bytes_taken As Byte
Private IsFinished As Boolean

' Consts
Private two_topmost_bits_values As Variant
Private pow2 As Variant
Private sha_k As Variant

Private Sub Class_Initialize()
    ' constants needed to simulate unsigned arithmetic
    two_topmost_bits_values = Array(0, &H40000000, &H80000000, &HC0000000)
    pow2 = Array(1&, 2&, 4&, 8&, 16&, 32&, 64&, 128&, 256&, 512&, 1024&, 2048&, 4096&, _
        8192&, 16384&, 32768, 65536, 131072, 262144, 524288, 1048576, 2097152, 4194304, _
        8388608, 16777216, 33554432, 67108864, 134217728, 268435456, 536870912, 1073741824, &H80000000)
    ' SHA-2 256 constants
    sha_k = Array( _
        &H428A2F98, &H71374491, &HB5C0FBCF, &HE9B5DBA5, &H3956C25B, &H59F111F1, &H923F82A4, &HAB1C5ED5, _
        &HD807AA98, &H12835B01, &H243185BE, &H550C7DC3, &H72BE5D74, &H80DEB1FE, &H9BDC06A7, &HC19BF174, _
        &HE49B69C1, &HEFBE4786, &HFC19DC6, &H240CA1CC, &H2DE92C6F, &H4A7484AA, &H5CB0A9DC, &H76F988DA, _
        &H983E5152, &HA831C66D, &HB00327C8, &HBF597FC7, &HC6E00BF3, &HD5A79147, &H6CA6351, &H14292967, _
        &H27B70A85, &H2E1B2138, &H4D2C6DFC, &H53380D13, &H650A7354, &H766A0ABB, &H81C2C92E, &H92722C85, _
        &HA2BFE8A1, &HA81A664B, &HC24B8B70, &HC76C51A3, &HD192E819, &HD6990624, &HF40E3585, &H106AA070, _
        &H19A4C116, &H1E376C08, &H2748774C, &H34B0BCB5, &H391C0CB3, &H4ED8AA4A, &H5B9CCA4F, &H682E6FF3, _
        &H748F82EE, &H78A5636F, &H84C87814, &H8CC70208, &H90BEFFFA, &HA4506CEB, &HBEF9A3F7, &HC67178F2 _
        )
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
    hash = Array(&H6A09E667, &HBB67AE85, &H3C6EF372, &HA54FF53A, &H510E527F, &H9B05688C, &H1F83D9AB, &H5BE0CD19)
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
    Dim sigma0 As Long
    Dim sigma1 As Long
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim e As Long
    Dim f As Long
    Dim g As Long
    Dim h As Long
    Dim ch_ As Long
    Dim t1_ As Long
    Dim t2_ As Long
    Dim maj_ As Long

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
        If idx > 15 Then
            sigma0 = Sigma0Expand(w(idx - 15))
            sigma1 = Sigma1Expand(w(idx - 2))
            w(idx) = UnsigSum(w(idx - 16), UnsigSum(sigma0, UnsigSum(w(idx - 7), sigma1)))
        End If
        sigma1 = Sigma1Normal(e)
        ch_ = (e And f) Xor ((Not e) And g)
        t1_ = UnsigSum(h, UnsigSum(sigma1, UnsigSum(ch_, UnsigSum(sha_k(idx), w(idx)))))
        sigma0 = Sigma0Normal(a)
        maj_ = (a And b) Xor (a And c) Xor (b And c)
        t2_ = UnsigSum(sigma0, maj_)
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

Friend Function RRot(ByVal x As Long, ByVal times As Byte) As Long
    times = times And 31
    If times = 0 Then
        RRot = x
        Exit Function
    End If
    Dim bit_31_shifted As Long
    Dim bit_to_be_31 As Long
    Dim lower_bits As Long
    bit_31_shifted = (x < 0&) And pow2(31 - times)
    x = x And &H7FFFFFFF
    bit_to_be_31 = x And pow2(times - 1)
    lower_bits = x \ pow2(times)
    RRot = (x - (lower_bits * pow2(times)) - bit_to_be_31) * pow2(32 - times) + lower_bits + bit_31_shifted
    If bit_to_be_31 Then RRot = RRot Or &H80000000
End Function

Friend Function RShift(ByVal x As Long, ByVal times As Byte) As Long
    times = times And 31
    If times = 0 Then
        RShift = x
        Exit Function
    End If
    RShift = (x And &H7FFFFFFF) \ pow2(times) + ((x < 0&) And pow2(31 - times))
End Function

Friend Function UnsigSum(ByVal x As Long, ByVal y As Long) As Long
    Dim topmost_bits As Long
    If x And &H80000000 Then topmost_bits = topmost_bits + 2&
    If x And &H40000000 Then topmost_bits = topmost_bits + 1&
    If y And &H80000000 Then topmost_bits = topmost_bits + 2&
    If y And &H40000000 Then topmost_bits = topmost_bits + 1&
    x = x And &H3FFFFFFF
    y = y And &H3FFFFFFF
    UnsigSum = x + y
    If UnsigSum And &H40000000 Then
        topmost_bits = topmost_bits + 1&
        UnsigSum = UnsigSum And &H3FFFFFFF
    End If
    UnsigSum = UnsigSum Or two_topmost_bits_values(topmost_bits And 3&)
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
        Case 2: ByteToLong = x * pow2(8)
        Case 1: ByteToLong = x * pow2(16)
        Case 0:
            If (x And &H80) Then
                ByteToLong = ((x And &H7F) * pow2(24)) Or &H80000000
            Else
                ByteToLong = x * pow2(24)
            End If
    End Select
End Function

Private Function FourBytesToLong(ByVal hhbyte As Byte, ByVal lhbyte As Byte, ByVal hlbyte As Byte, ByVal llbyte As Byte) As Long
    Dim lower As Long
    lower = (lhbyte * pow2(16)) Or (hlbyte * pow2(8)) Or llbyte
    If (hhbyte And &H80) Then
        FourBytesToLong = ((hhbyte And &H7F) * pow2(24)) Or &H80000000 Or lower
    Else
        FourBytesToLong = hhbyte * pow2(24) Or lower
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

