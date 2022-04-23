Attribute VB_Name = "JSON"
'@IgnoreModule ModuleWithoutFolder, ProcedureCanBeWrittenAsFunction, ObsoleteLetStatement, ParameterCanBeByVal, FunctionReturnValueDiscarded

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '
'
'                                        JSON
'
' JSON decoder/encoder.
'
' For documentation, licensing, updates etc. see:
' <https://github.com/kimbar/VBA-Native-Tools>
'
' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

Option Explicit

Private Enum JSONState
    json_s_JB = 0        ' JSON Begin
    json_s_AIB = 1       ' Array Item Begin
    json_s_OKB = 2       ' Object Key Begin
    json_s_OVB = 3       ' Object Value Begin
    json_s_JE = 4        ' JSON End
    json_s_AIE = 5       ' Array Item End
    json_s_OKE = 6       ' Object Key End
    json_s_OVE = 7       ' Object Value End
    json_s_SPECIALS = 8  ' Vitrual state - delimiter for special states
    json_s_CE_ = 9       ' Container (Array or Object) End
    json_s_CPE_ = 10     ' Container Possibly Empty
    json_s_END_ = 11     ' Successful end of input
    json_s_ERRORS = 12   ' Vitrual state - delimiter for error states
    json_s_JBERR = 13
    json_s_AIBERR = 14
    json_s_OKBERR = 15
    json_s_OVBERR = 16
    json_s_JEERR = 17
    json_s_AIEERR = 18
    json_s_OKEERR = 19
    json_s_OVEERR = 20
End Enum

Private Enum JSONToken
    json_t_s = 0         ' String - starts with `"`
    json_t_p = 1         ' Primitive - bool, number or null - starts with one of `nft-0123456789`
    json_t_ob = 2        ' object begin - consists of `{`
    json_t_ab = 3        ' array begin - consists of `[`
    json_t_oe = 4        ' object end - consists of `}`
    json_t_ae = 5        ' array end - consists of `]`
    json_t_is = 6        ' item separator - consists of `,`
    json_t_kvs = 7       ' key-value separator - consists of `:`
    json_t_eos = 8       ' end of stream - when cursor is greater than length od string
    json_t_ill = 9       ' illegal token - any other
    json_t_TOTAL = 10    ' virtual token - total number of tokens
End Enum

Const state_change_map As String = "EECBNNNNNNFFCBOKOOOOGPPPKPPPPPHHCBQQQQQQRRRRRRRRLRSSSSSJBSSSTTTTTTTDTTUUUUJUCUUU"

' Customizable block
' ------------------

' functions to perform elementary tasks during JSON parsing

' Customize creation of new JSON object i.e. `{}`
Private Function ObjectNew() As Variant
    Set ObjectNew = CreateObject("Scripting.Dictionary")
End Function

' Customize adding new item to a JSON object i.e. `{"key": item}`
Private Sub ObjectAppend(ByRef obj As Variant, ByRef key As String, ByRef item As Variant)
    obj.Add key:=key, item:=item
End Sub

' Customize postprocessing of a finished JSON object. No `ObjectAppend` will be called after this point
' The result of this function may be arbitrary value / VBA object
Private Function ObjectPostprocess(ByRef obj As Variant) As Variant
    Set ObjectPostprocess = obj
End Function

' Customize creation of new JSON array i.e. `[]`
Private Function ArrayNew() As Variant
    Set ArrayNew = New Collection
End Function

' Customize adding new item to a JSON array i.e. `[item]`, `index` doesn't have to be used
'@Ignore ParameterNotUsed
Private Sub ArrayAppend(ByRef arr As Variant, ByRef index As Long, ByRef item As Variant)
    arr.Add item
End Sub

' Customize postprocessing of a finished JSON array. No `ArrayAppend` will be called after this point
' The result of this function may be arbitrary value / VBA object
Private Function ArrayPostprocess(ByRef arr As Variant) As Variant
    Set ArrayPostprocess = arr
End Function

'functions to perform elementary tasks during building of JSON

' Customize preprocesing of data. `PreProcess*` must return string, number, bool, `Null`, Collection, Array or
' Scripting.Dictionary. All keys in Scripting.Dictionary must be strings.
' `PreprocessObject` takes an object as input, but can return value and vice versa.
Private Function PreProcessObject(ByRef obj As Variant) As Variant
    Set PreProcessObject = obj
End Function

Private Function PreProcessValue(ByRef val As Variant) As Variant
    Let PreProcessValue = val
End Function

' Customize diferentiation between integers and floating-points
Private Function EncodeNumber(ByRef num As Variant) As String
    EncodeNumber = CStr(num)
End Function

'@EntryPoint
Public Function Decode(ByRef stream As String) As Variant

    Decode = Null

    Dim token As JSONToken
    Dim state As JSONState: Let state = json_s_JB
    Dim cursor As Long: Let cursor = 1
    Dim vstack As Variant
    Dim astack As Variant
    Dim temporary As Variant
    Dim tokenOK As Boolean
    Dim error_message As String

    Do
        ' Operations based on the state
        Select Case state
            Case json_s_JB
                Set vstack = NewStack
                Set astack = NewStack
            Case json_s_AIE
                StackPop vstack, saveTo:=temporary
                ArrayAppend StackTop(vstack), StackTop(astack), temporary
            Case json_s_OVE
                StackPop vstack, saveTo:=temporary
                ObjectAppend StackTop(vstack), StackTop(astack), temporary
            Case json_s_OKE
                StackPop astack ' remove `Null`
                StackPush astack, StackPop(vstack)
            Case Else
                ' NOP
        End Select

        token = DetectToken(stream, cursor)
        tokenOK = True
        ' Operations based on encountered token
        Select Case token
            Case json_t_s
                StackPush vstack, ConsumeString(stream, cursor, tokenOK)
            Case json_t_p
                StackPush vstack, ConsumePrimitive(stream, cursor, tokenOK)
            Case json_t_ob
                StackPush vstack, ObjectNew()
                StackPush astack, 0    ' Empty Object has its first address set to `0`, just like an Array
            Case json_t_ab
                StackPush vstack, ArrayNew()
                StackPush astack, 0
            Case json_t_oe
                If Not IsStackEmpty(vstack) Then _
                    StackPush vstack, ObjectPostprocess(StackPop(vstack))
            Case json_t_ae
                If Not IsStackEmpty(vstack) Then _
                    StackPush vstack, ArrayPostprocess(StackPop(vstack))
            Case json_t_is
                ' Advance current address
                If Not IsStackEmpty(astack) Then
                    ' if `astack` is empty it will be an error soon
                    temporary = StackPop(astack)
                    If TypeName(temporary) = "String" Then
                        ' Temporally set the current address to `Null`, until we read a propper key
                        StackPush astack, Null
                    Else
                        StackPush astack, temporary + 1
                    End If
                End If
            '@Ignore EmptyCaseBlock
            Case json_t_kvs
                ' NOP
            Case json_t_eos
                ' Save the value as an output
                If Not IsStackEmpty(vstack) Then
                ' if `vstack` is empty it will be an error soon
                    StackPop vstack, saveTo:=Decode
                End If
            '@Ignore EmptyCaseBlock
            Case json_t_ill
                ' NOP - ERROR
        End Select
        cursor = cursor + 1
        ' The `DetectToken` assumes the token just by the first character, then apropriate `Consume*` functions
        ' try to parse it completely. If they fail they set `tokenOK=False`
        If Not tokenOK Then token = json_t_ill

        ' calculate new state with `state_change_map` - this is basically a hack to implement const array
        state = Asc(Mid$(state_change_map, state * json_t_TOTAL + token + 1, 1)) - Asc("A")

        ' Special states
        ' They are not proper states of the machine - they are recalculated to proper states based on the contents
        ' of value stack and address stack. It is done to avoid redundancy of information and proliferation of states.
        If (state > json_s_SPECIALS) And (state < json_s_ERRORS) Then
            If state = json_s_CPE_ Then
                If StackTop(astack) = 0 Then  ' The address is `0` iff no item has been added yet, even if it is an Object
                                              ' Objects have addresses set to `Null` between items, but to `0` at the beggining
                    state = json_s_CE_
                Else
                    ' Here we have a "trailing comma error", that is a `{... ,}` or `[... ,]` case
                    If TypeName(StackTop(astack)) = "String" Then
                        state = json_s_OKBERR
                    Else
                        state = json_s_AIBERR
                    End If
                End If
            End If
            If state = json_s_CE_ Then
                StackPop astack
                If IsStackEmpty(astack) Then
                    ' If it was last container in the stack we expect the end of the JSON now
                    state = json_s_JE
                Else
                    ' If it was not the last container we need to retrieve what kind of container has just ended
                    ' from the type of the top value on the address stack
                    If TypeName(StackTop(astack)) = "String" Then
                        state = json_s_OVE
                    Else
                        state = json_s_AIE
                    End If
                End If
            End If
            If state = json_s_END_ Then
                ' Propper end of the JSON has been reached
                Exit Do
            End If
        End If

        ' Error states
        If state > json_s_ERRORS Then
            error_message = _
                "JSON parsing error (" & (state - json_s_ERRORS - 1) * json_t_TOTAL + token & "): at " & _
                Array( _
                    "begining of JSON", "begining of array item", "begining of object key", "begining of object value", _
                    "end of JSON", "end of array item", "end of object key", "end of object value" _
                )(state - json_s_ERRORS - 1) _
                & " an unexpected " & _
                Array( _
                    "string start", "primitive (bool, number or null)", "object start `{`", "array start `[`", "object end `}`", _
                    "array end `]`", "item separator `,`", "key-value separator `:`", "end of stream", "illegal character" _
                )(token) & _
                " has been found at position " & cursor - 1
            Err.Raise 321, Description:=error_message
            Exit Do
        End If
    Loop

End Function

Private Sub SkipWhiteSpaces(ByRef stream As String, ByRef cursor As Long)
    Dim code As Long
    Do
        If cursor > Len(stream) Then Exit Sub
        code = AscW(Mid$(stream, cursor, 1))
        If (code <> 32) And (code <> 9) And (code <> 10) And (code <> 13) Then Exit Sub
        cursor = cursor + 1
    Loop
End Sub

Private Function DetectToken(ByRef stream As String, ByRef cursor As Long) As JSONToken
    SkipWhiteSpaces stream, cursor
    If cursor > Len(stream) Then
        DetectToken = json_t_eos
    Else
        Dim char As String: Let char = Mid$(stream, cursor, 1)
        If char = """" Then
            DetectToken = json_t_s
        ElseIf IsNumeric(char) Or (char = "-") Or (char = "n") Or (char = "f") Or (char = "t") Then
            DetectToken = json_t_p
        ElseIf char = "{" Then: DetectToken = json_t_ob
        ElseIf char = "[" Then: DetectToken = json_t_ab
        ElseIf char = "}" Then: DetectToken = json_t_oe
        ElseIf char = "]" Then: DetectToken = json_t_ae
        ElseIf char = "," Then: DetectToken = json_t_is
        ElseIf char = ":" Then: DetectToken = json_t_kvs
        Else: DetectToken = json_t_ill
        End If
    End If
End Function

Private Function ConsumePrimitive(ByRef stream As String, ByRef cursor As Long, ByRef IsOK As Boolean) As Variant
    IsOK = True
    If Mid$(stream, cursor, 4) = "true" Then
        ConsumePrimitive = True
        cursor = cursor + 3
    ElseIf Mid$(stream, cursor, 5) = "false" Then
        ConsumePrimitive = False
        cursor = cursor + 4
    ElseIf Mid$(stream, cursor, 4) = "null" Then
        ConsumePrimitive = Null
        cursor = cursor + 3
    Else
        ConsumePrimitive = ConsumeNumber(stream, cursor, IsOK)
    End If
End Function

Private Function ConsumeNumber(ByRef stream As String, ByRef cursor As Long, Optional ByRef IsOK As Boolean) As Variant
    Dim code As Long

    code = AscW(Mid$(stream, cursor, 1))
    If (code <> AscW("-")) And ((code < AscW("0")) Or (code > AscW("9"))) Then
        IsOK = False
        Exit Function
    End If

    Dim char As String
    Dim exp_has_sign As Long
    Dim istart As Long
    Dim sign_part As String
    Dim int_part As String
    Dim frac_part As String
    Dim exp_part As String

    If Mid$(stream, cursor, 1) = "-" Then
        sign_part = "-"
        cursor = cursor + 1
    End If

    istart = cursor
    If Mid$(stream, cursor, 1) = "0" Then
        cursor = cursor + 1
    Else
        ConsumeDigits stream, cursor
    End If

    int_part = Mid$(stream, istart, cursor - istart)

    char = Mid$(stream, cursor, 1)
    If char = "." Then GoTo do_frac
    If (char = "e") Or (char = "E") Then GoTo do_exp
    GoTo finish

do_frac:
    istart = cursor
    cursor = cursor + 1
    ConsumeDigits stream, cursor
    frac_part = Mid$(stream, istart, cursor - istart)

    char = Mid$(stream, cursor, 1)
    If (char = "e") Or (char = "E") Then GoTo do_exp
    GoTo finish

do_exp:
    istart = cursor
    cursor = cursor + 1
    char = Mid$(stream, cursor, 1)
    If (char = "+") Or (char = "-") Then
        exp_has_sign = 1
        cursor = cursor + 1
    End If
    ConsumeDigits stream, cursor
    exp_part = Mid$(stream, istart, cursor - istart)

finish:
    If (Len(int_part) = 0) Or (Len(frac_part) = 1) Or (Len(exp_part) - exp_has_sign = 1) Then
        IsOK = False
        Exit Function
    End If
    int_part = sign_part & int_part

    If (Len(frac_part) = 0) And (Len(exp_part) = 0) Then
        ConsumeNumber = CLng(int_part)
    Else
        ConsumeNumber = CDbl(int_part & frac_part & exp_part)
    End If
    cursor = cursor - 1
    IsOK = True
End Function

Private Sub ConsumeDigits(ByVal stream As String, ByRef cursor As Long)
    Dim code As Long
    Do
        If cursor > Len(stream) Then
            Exit Do
        End If
        code = AscW(Mid$(stream, cursor, 1))
        If (code < AscW("0")) Or (code > AscW("9")) Then
            Exit Do
        Else
            cursor = cursor + 1
        End If
    Loop
End Sub

Private Function ConsumeString(ByRef stream As String, ByRef cursor As Long, Optional ByRef IsOK As Boolean) As String
    IsOK = True

    Dim quote_position As Long: Let quote_position = cursor
    Dim carret As Long

    Do
        quote_position = InStr(quote_position + 1, stream, """")
        If quote_position = 0 Then
            IsOK = False
            Exit Function
        End If
        If Mid$(stream, quote_position - 1, 1) <> "\" Then
            Exit Do
        End If
    Loop

    ConsumeString = DecodeStringPriv(Mid$(stream, cursor + 1, quote_position - cursor - 1), carret, IsOK)
    cursor = cursor + carret
End Function

'@EntryPoint
Public Function DecodeString(ByVal stream As String) As String
    Dim IsOK As Boolean
    Dim cursor As Long: Let cursor = 1
    DecodeString = DecodeStringPriv(stream, cursor, IsOK)
    If Not IsOK Then Err.Raise 321, Description:="Unable to decode string at position " & CStr(cursor)
End Function

Private Function DecodeStringPriv(ByVal stream As String, ByRef cursor As Long, ByRef IsOK As Boolean) As String
    ' This function is sligtly more permisive than the standard allows - it allows for
    ' &H00 - &H1F characters to appear in the string. That includes "real" (unescaped) tabs, LFs and CRs
    cursor = 1
    IsOK = True
    DecodeStringPriv = vbNullString

    Dim rsi As Long

    Do
        If cursor > Len(stream) Then Exit Do
        rsi = InStr(cursor, stream, "\")
        If rsi = 0 Then
            ' No more escape sequences
            DecodeStringPriv = DecodeStringPriv & Mid$(stream, cursor)
            cursor = Len(stream) + 1
            Exit Do
        End If
        If rsi = Len(stream) Then
            ' Solitary escape sequence marker `\` at the end of string is not possible
            ' This code is unreachable in JSON, because a `\"` sequence wouldn't be recognised as a string end
            IsOK = False
            Exit Do
        End If
        ' Consume everything up to escape sequence
        DecodeStringPriv = DecodeStringPriv & Mid$(stream, cursor, rsi - cursor)
        ' All (most) sequences are at least 2 chars long
        cursor = rsi + 2
        Select Case Mid$(stream, cursor - 1, 1)
            Case """": DecodeStringPriv = DecodeStringPriv & """"
            Case "\": DecodeStringPriv = DecodeStringPriv & "\"
            Case "/": DecodeStringPriv = DecodeStringPriv & "/"
            Case "b": DecodeStringPriv = DecodeStringPriv & ChrW$(&H8)
            Case "f": DecodeStringPriv = DecodeStringPriv & ChrW$(&HC)
            Case "n": DecodeStringPriv = DecodeStringPriv & ChrW$(&HA)
            Case "r": DecodeStringPriv = DecodeStringPriv & ChrW$(&HD)
            Case "t": DecodeStringPriv = DecodeStringPriv & ChrW$(&H9)
            Case "u"
                ' We need 4 more chars to exist
                If cursor + 3 > Len(stream) Then
                    IsOK = False
                    Exit Do
                End If
                On Error GoTo catch_bad_hex
                DecodeStringPriv = DecodeStringPriv & ChrW$(CLng("&H" & Mid$(stream, cursor, 4)))
                On Error GoTo 0
                ' `\u0000` sequence is 6 chars long, so we need additional 4
                cursor = cursor + 4
            Case Else
                ' Illegal escape sequence
                IsOK = False
                Exit Do
        End Select
    Loop
    Exit Function
catch_bad_hex:
    IsOK = False
    Err.Clear
End Function

'@EntryPoint
Public Function Encode(ByRef data As Variant) As String

    Dim vstack As Variant: Set vstack = NewStack
    Dim astack As Variant: Set astack = NewStack
    Dim temporary As Variant

    StackPush vstack, data

    Do
        If IsStackEmpty(vstack) Then Exit Do

        StackPop vstack, saveTo:=temporary
        If IsObject(temporary) Then
            StackPush vstack, PreProcessObject(temporary)
        Else
            StackPush vstack, PreProcessValue(temporary)
        End If

        Select Case BasicTypesCast(TypeName(StackTop(vstack)))
            Case "Dictionary": Encode = Encode & BuildJSONContainer(astack, vstack, "{}")
            Case "JSON_Array": Encode = Encode & BuildJSONContainer(astack, vstack, "[]")
            Case "String": Encode = Encode & """" & EncodeString(StackPop(vstack)) & """"
            Case "JSON_Number": Encode = Encode & EncodeNumber(StackPop(vstack))
            Case "Boolean"
                If StackPop(vstack) Then
                    Encode = Encode & "true"
                Else
                    Encode = Encode & "false"
                End If
            Case "Null"
                Encode = Encode & "null"
                StackPop vstack
            Case Else
                Err.Raise 321, Description:="Bad type of data member"
        End Select
    Loop
End Function

Private Function BasicTypesCast(ByRef tname As String) As String
    BasicTypesCast = tname
    If (tname = "Double") Or (tname = "Integer") Then BasicTypesCast = "JSON_Number"
    If (tname = "Collection") Or (tname = "Variant()") Then BasicTypesCast = "JSON_Array"
End Function

Private Function GetIndexDescriptor(ByRef Container As Variant) As Variant
    Select Case TypeName(Container)
        Case "Dictionary"
            Set GetIndexDescriptor = New Collection
            Dim key As Variant
            ' This `Null` marks the begining of the Object key-set
            GetIndexDescriptor.Add Null
            For Each key In Container.Keys()
                If TypeName(key) = "String" Then
                    GetIndexDescriptor.Add key
                Else
                    Err.Raise 321, Description:="Bad type of Object key"
                End If
            Next
        Case "Collection"
            Let GetIndexDescriptor = Array(1, Container.Count, 1)
        Case "Variant()"
            ' The meaning of this values is: current index, last index, first index
            Let GetIndexDescriptor = Array(LBound(Container), UBound(Container), LBound(Container))
        Case Else
            Err.Raise 321, Description:="Bad type of Object"
    End Select
End Function

Private Function AdvanceIndexDescriptor(ByRef astack As Variant, ByRef vstack As Variant, ByRef is_first As Boolean) As Variant
    Dim idx_descriptor As Variant
    is_first = False

    Select Case TypeName(StackTop(astack))
        Case "Collection"
            Set idx_descriptor = StackPop(astack)
            If idx_descriptor.Count > 0 Then
                If IsNull(idx_descriptor.item(1)) Then
                    ' If key is equal to `Null` it means that we are at the begining of the object
                    ' and this `Null` should be discarded
                    is_first = True
                    idx_descriptor.Remove 1
                End If
            End If
            If idx_descriptor.Count = 0 Then
                AdvanceIndexDescriptor = Null
            Else
                AdvanceIndexDescriptor = idx_descriptor.item(1)
                idx_descriptor.Remove 1
            End If
        Case "Variant()"
            Let idx_descriptor = StackPop(astack)
            If idx_descriptor(0) = idx_descriptor(2) Then is_first = True
            If idx_descriptor(0) > idx_descriptor(1) Then
                AdvanceIndexDescriptor = Null
            Else
                AdvanceIndexDescriptor = idx_descriptor(0)
                idx_descriptor(0) = idx_descriptor(0) + 1
            End If
        Case Else
            ' internal error - `GetIndexDescriptor` produced something unrecognizable
    End Select
    If IsNull(AdvanceIndexDescriptor) Then
        StackPop vstack
    Else
        StackPush astack, idx_descriptor
    End If
End Function

Private Function BuildJSONContainer(ByRef astack As Variant, ByRef vstack As Variant, ByRef braces As String) As String
    Dim key As Variant
    Dim is_first As Boolean

    If StackSize(astack) < StackSize(vstack) Then
        StackPush astack, GetIndexDescriptor(StackTop(vstack))
        BuildJSONContainer = Left$(braces, 1)    ' opening `{` or `[`
    Else
        key = AdvanceIndexDescriptor(astack, vstack, is_first)
        If IsNull(key) Then
            BuildJSONContainer = Right$(braces, 1)    ' closing `}` or `]`
        Else
            If Not is_first Then BuildJSONContainer = ","
            If braces = "{}" Then BuildJSONContainer = BuildJSONContainer & """" & EncodeString(CStr(key)) & """:"
            StackPush vstack, StackTop(vstack)(key)
        End If
    End If
End Function

Public Function EncodeString(ByRef stream As String) As String
    Dim cursor As Long
    Dim code As Long
    For cursor = 1 To Len(stream)
        code = AscW(Mid$(stream, cursor, 1))
        Select Case code
            Case &H22: EncodeString = EncodeString & "\"""
            Case &H5C: EncodeString = EncodeString & "\\"
            Case &H8: EncodeString = EncodeString & "\b"
            Case &HC: EncodeString = EncodeString & "\f"
            Case &HA: EncodeString = EncodeString & "\n"
            Case &HD: EncodeString = EncodeString & "\r"
            Case &H9: EncodeString = EncodeString & "\t"
            Case Else
                If (code < &H20) Or (code > &H24F) Or ((code >= &H80) And (code <= &H7F)) Or (code = &HAD) Then
                    EncodeString = EncodeString & "\u" & Left$("000", 4 - Len(Hex$(code))) & Hex$(code)
                Else
                    EncodeString = EncodeString & ChrW$(code)
                End If
        End Select
    Next
End Function

' General stack implementation
' ----------------------------

Private Function NewStack() As Variant
    Set NewStack = New Collection
End Function

Private Sub StackPush(ByRef stack As Variant, ByRef item As Variant)
    stack.Add item
End Sub

Private Function StackPop(ByRef stack As Variant, Optional ByRef saveTo As Variant) As Variant
    If IsObject(stack(stack.Count)) Then
        Set saveTo = stack(stack.Count)
        Set StackPop = saveTo
    Else
        saveTo = stack(stack.Count)
        StackPop = saveTo
    End If
    stack.Remove stack.Count
End Function

Private Function StackTop(ByRef stack As Variant) As Variant
    If IsObject(stack(stack.Count)) Then
        Set StackTop = stack(stack.Count)
    Else
        StackTop = stack(stack.Count)
    End If
End Function

Private Function IsStackEmpty(ByRef stack As Variant) As Boolean
    IsStackEmpty = (stack.Count = 0)
End Function

Private Function StackSize(ByRef stack As Variant) As Long
    StackSize = stack.Count
End Function


