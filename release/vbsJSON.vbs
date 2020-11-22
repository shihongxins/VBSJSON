' 
' Author: shihongxins
' Date: 2020-11-21
' GitHub: https://github.com/shihongxins
' LICENSE: MIT License https://github.com/shihongxins/vbsJSON/blob/main/LICENSE
' 
Option Explicit
Class vbsJSON
    Private WhiteSpace, NumberRegExp
    Private ParserError

    Private Function ClearParserError()
        ParserError = ""
    End Function

    Private Function SetParserError(ByVal Discription)
        ParserError = CStr(Discription)
        Err.Raise 6,,ParserError
    End Function

    Public Function GetParserError()
        GetParserError = ParserError
    End Function

    Private Sub Class_Initialize
        Call ClearParserError
        Whitespace = " " & vbTab & vbCr & vbLf

        Set NumberRegExp = New RegExp
        NumberRegExp.Pattern = "^([\+\-\.]?(0|[1-9]\d*))(\.\d+)?([Ee\+\-\.\d+])?$"
        NumberRegExp.Global = False
        NumberRegExp.MultiLine = True
        NumberRegExp.IgnoreCase = True
    End Sub

    Private Function EncodeUnicode(ByVal str)
        Dim i,uChrCode,uChr,uStr
        For i=1 To Len(str)
            uChrCode = AscW(Mid(str,i,1))
            Select Case uChrCode
                Case 8      uChr = "\b"
                Case 9      uChr = "\t"
                Case 10     uChr = "\n"
                Case 12     uChr = "\f"
                Case 13     uChr = "\r"
                Case 34     uChr = "\"""
                Case 39     uChr = "\'"
                Case 92     uChr = "\\"
                Case Else
                    If uChrCode<32 Or uChrCode>127 Then
                        uChr = "\u" & Right("0000" & Hex(uChrCode), 4)
                    Else
                        uChr = ChrW(uChrCode)
                    End If
            End Select
            uStr = uStr & uChr
        Next
        EncodeUnicode = uStr
    End Function

    Private Function DecodeUnicode(ByVal str)
        Dim i,uChr,uStr
        For i=1 To Len(str)
            uChr = Mid(str,i,1)
            If uChr="\" Then
                i = i + 1
                uChr = Mid(str,i,1)
                Select Case uChr
                    Case "b"    uChr = ChrW(8)
                    Case "t"    uChr = ChrW(9)
                    Case "n"    uChr = ChrW(10)
                    Case "f"    uChr = ChrW(12)
                    Case "r"    uChr = ChrW(13)
                    Case """"   uChr = ChrW(34)
                    Case "'"    uChr = ChrW(39)
                    Case "/"    uChr = ChrW(47)
                    Case "\"    uChr = ChrW(92)
                    Case "u"
                        If i+4 <=Len(str) Then
                            uChr = ChrW("&H" & Mid(str,i + 1,4))
                            i = i + 4
                        Else
                            uChr = "\" & uChr
                        End If
                End Select
            End If
            uStr = uStr & uChr
        Next
        DecodeUnicode = uStr
    End Function

    Private Function skipWhiteSpace(ByRef str,ByRef index)
        Do While index>0 And index<=Len(str)
            If InStr(WhiteSpace,Mid(str,index,1))>0 Then
                index = index + 1
            Else
                Exit Do
            End If    
        Loop
        skipWhiteSpace = index
    End Function

    Public Function stringify(ByRef obj)
        Call ClearParserError
        Dim buf:Set buf = CreateObject("Scripting.Dictionary")
        Dim g
        Select Case VarType(obj)
            Case vbNull
                buf.Add buf.Count, "null"
            Case vbBoolean
                If obj Then
                    buf.Add buf.Count, "true"
                Else
                    buf.Add buf.Count, "false"
                End If
            Case vbInteger, vbLong, vbSingle, vbDouble
                buf.Add buf.Count, obj
            Case vbString
                buf.Add buf.Count, """" & EncodeUnicode(obj) & """"
            Case vbArray, vbVariant, vbArray+vbVariant
                g = True
                Dim Value
                buf.Add buf.Count, "["
                For Each Value In obj
                    If g Then
                        g = False
                    Else
                        buf.Add buf.Count, ","
                    End If
                    buf.Add buf.Count, stringify(Value)
                Next
                buf.Add buf.Count, "]"
            Case vbObject
                If TypeName(obj) = "Dictionary" Then
                    g = True
                    Dim Key
                    buf.Add buf.Count, "{"
                    For Each Key In obj.Keys
                        If g Then
                            g = False
                        Else
                            buf.Add buf.Count, ","
                        End If
                        buf.Add buf.Count, """" & EncodeUnicode(Key) & """" & ":" & stringify(obj.Item(Key))
                    Next
                    buf.Add buf.Count, "}"
                Else
                    Call SetParserError("None dictionary in object!")
                End If
            Case Else
                buf.Add buf.Count, """" & EncodeUnicode(CStr(obj)) & """"
        End Select
        stringify = Join(buf.Items, "")
    End Function

    Public Function parse(ByRef str)
        Call ClearParserError
        Select Case VarType(str)
            Case vbNull
                parse = Null
                Exit Function
            Case vbInteger, vbLong, vbSingle, vbDouble
                parse = str
                Exit Function
            Case vbBoolean
                parse = str
                Exit Function
            Case vbString
            Case Else
                Call SetParserError("Uncaught SyntaxError: Invalid JSON input,Unknown Type")
        End Select

        str = CStr(str)
        If Trim(str)="" Then
            Call SetParserError("Uncaught SyntaxError: Unexpected end of JSON input,Empty String")
        End If
        Dim FirstChar
        FirstChar = Mid(str,skipWhiteSpace(str,1),1)
        If FirstChar="{" Then
            Set parse = parseInit(str,1)
        Else
            parse = parseInit(str,1)
        End If
    End Function

    Private Function parseInit(ByRef str,ByRef index)
        index = skipWhiteSpace(str,index)
        Select Case Mid(str,index,1)
            Case "{"
                Set parseInit = parseObject(str,index)
            Case "["
                parseInit = parseArray(str,index)
            Case Else
                parseInit = parseBase(str,index)
        End Select
        index = index + 1
        index = skipWhiteSpace(str,index)
        If index <= Len(str) And Mid(str,index,1)<>"" Then
            Call SetParserError("Uncaught SyntaxError: Unexpected token " & Mid(str,index,1) & " in JSON at position " & index)
        End If
    End Function

    Private Function parseBase(ByVal str,ByVal index)
        Dim Char
        Char = Mid(str,index,1)
        Select Case Char
            Case "n"
                parseBase = parseNull(str,index)
            Case "t", "f"
                parseBase = parseBoolean(str,index)
            Case Else
                If InStr("-0123456789.", Char) Then
                    parseBase = parseNumber(str,index)
                Else
                    Call SetParserError("Uncaught SyntaxError: Unexpected token " & Char & " in JSON at position " & CStr(index))
                End If
        End Select
    End Function

    Private Function findFirstCharNotInString(ByVal testString,ByVal validType)
        Dim i, index, Char
        index = skipWhiteSpace(testString,1)
        Char = ""
        For i = index To Len(testString)
            If InStr(validString,Mid(testString,i,1))>0 Then
                'Continue
            Else
                Char = Mid(testString,i,1)
                Call SetParserError("Uncaught SyntaxError: Unexpected token " & Char & " in JSON at position " & CStr(i))
            End If
        Next
        
        If Char="" And validString="+-0123456789.Ee" Then
            Call SetParserError("Uncaught SyntaxError: Unexpected number in JSON of " & testString)
        End If
    End Function

    Private Function parseObject(ByRef str,ByRef index)
        Dim Char, Quote, Key, KeyEndFlag, Value
        Char = ""
        Key = ""
        Value = Empty
        If Mid(str,index,1)<>"{" Then
            Call SetParserError("Invalid Object at position " & index & " : " & Mid(str, index) & vbCrLf)
        End If

        Dim obj:Set obj = CreateObject("Scripting.Dictionary")
        KeyEndFlag = True
        Do
            index = index + 1
            If index > Len(str) Then
                Call SetParserError("Uncaught SyntaxError: Unexpected end of JSON input,Missing '}'")
            End If
            If KeyEndFlag=True Then
                index = skipWhiteSpace(str,index)
                Char = Mid(str,index,1)
                If Char="""" Or Char="'" Then
                    Quote = Char
                    KeyEndFlag = False
                ElseIf Char="," Then
                    If VarType(Value)=vbEmpty Then
                        Call SetParserError("Uncaught SyntaxError: Unexpected token " & Char & " in JSON at position " & CStr(index))
                    Else
                        Key = ""
                        Value = Empty
                    End If
                ElseIf Char="}" Then
                    Exit DO
                Else
                    Call SetParserError("Uncaught SyntaxError: Unexpected token " & Char & " in JSON at position " & CStr(index))
                End If
            Else
                Char = Mid(str,index,1)
                If Char<>Quote Then
                    Key = Key + Char
                ElseIf Char=Quote Then
                    KeyEndFlag = True
                    index = index + 1
                    index = skipWhiteSpace(str,index)
                    Char = Mid(str,index,1)
                    If Char=":" Then
                        index = index + 1
                        index = skipWhiteSpace(str,index)
                        Dim FirstChar
                        FirstChar = Mid(str,index,1)
                        If FirstChar="{" Then
                            Set Value = parseValue(str,index)
                        Else
                            Value = parseValue(str,index)
                        End If
                        Key = DecodeUnicode(Key)
                        If obj.Exists(Key) Then
                            obj.Remove(Key)
                        End If
                        obj.Add Key, Value
                    Else
                        Call SetParserError("Uncaught SyntaxError: Unexpected token " & Char & " in JSON at position " & CStr(index) & "Expecting ':' delimiter")
                    End If
                End If
            End If
        Loop
        Set parseObject = obj
    End Function

    Private Function parseArray(ByRef str,ByRef index)
        Dim Char, Value
        Char = ""
        Value = Empty
        If Mid(str,index,1)<>"[" Then
            Call SetParserError("Invalid Array at position " & index & " : " & Mid(str, index) & vbCrLf)
        End If

        Dim obj:Set obj = CreateObject("Scripting.Dictionary")
        index = index + 1
        Do
            If VarType(Value)=vbEmpty Then
                index = skipWhiteSpace(str,index)
            End If
            If index > Len(str) Then
                Call SetParserError("Uncaught SyntaxError: Unexpected end of JSON input,Missing ']'")
            End If
            Char = Mid(str,index,1)
            If Char="," Then
                If VarType(Value)=vbEmpty Then
                    Call SetParserError("Uncaught SyntaxError: Unexpected token " & Char & " in JSON at position " & CStr(index))
                Else
                    Value = Empty
                End If
            ElseIf Char="]" Then
                Exit Do
            Else
                Dim FirstChar
                FirstChar = Mid(str,index,1)
                If FirstChar="{" Then
                    Set Value = parseValue(str,index)
                Else
                    Value = parseValue(str,index)
                End If
                obj.Add obj.Count, Value
            End If
            index = index + 1
            If VarType(Value)<>vbEmpty Then
                index = skipWhiteSpace(str,index)
            End If
        Loop
        parseArray = obj.Items
    End Function

    Private Function parseValue(ByRef str,ByRef index)
        Dim Char
        index = skipWhiteSpace(str,index)
        Char = Mid(str,index,1)
        Select Case Char
            Case "{"
                Set parseValue = parseObject(str,index)
            Case "["
                parseValue = parseArray(str,index)
            Case "n"
                parseValue = parseNull(str,index)
            Case "t", "f"
                parseValue = parseBoolean(str,index)
            Case """", "'"
                parseValue = parseString(str,index)
            Case Else 
                If InStr("-0123456789.", Char) Then
                    parseValue = parseNumber(str,index)
                Else
                    Call SetParserError("Uncaught SyntaxError: Unexpected token " & Char & " in JSON at position " & CStr(index))
                End If
        End Select
    End Function

    Private Function parseNull(ByRef str, ByRef index)
        If index+4 <= Len(str) And Mid(str,index,4) = "null" Then
            parseNull = Null
            index = index + 3
        Else
            Dim Char
            Char = Mid(str,index,1)
            If Char="n" Then
                If index+4 <= Len(str) Then
                    Call findFirstCharNotInString(Mid(str,index,4), "null")
                Else
                    Call SetParserError("Uncaught SyntaxError: Unexpected end of JSON input,parse null of " & Mid(str,index,Len(str)-index) & " at " & index)
                End If
            Else
                Call SetParserError("Uncaught SyntaxError: Unexpected null in JSON of " & Char & " at " & index)
            End If
        End If
    End Function

    Private Function parseBoolean(ByRef str, ByRef index)
        If index+5 <= Len(str) And Mid(str,index,5) = "false" Then
            parseBoolean = False
            index = index + 4
        ElseIf index+4 <= Len(str) And Mid(str,index,4) = "true" Then
            parseBoolean = True
            index = index + 3
        Else
            Dim Char
            Char = Mid(str,index,1)
            If Char="f" Then
                If index+5 <= Len(str) Then
                    Call findFirstCharNotInString(Mid(str,index,5), "false")
                Else
                    Call SetParserError("Uncaught SyntaxError: Unexpected end of JSON input,parse boolean of " & Mid(str,index,Len(str)-index) & " at " & index)
                End If
            ElseIf Char="t" Then
                If index+4 <= Len(str) Then
                    Call findFirstCharNotInString(Mid(str,index,4), "true")
                Else
                    Call SetParserError("Uncaught SyntaxError: Unexpected end of JSON input,parse boolean of " & Mid(str,index,Len(str)-index) & " at " & index)
                End If
            Else
                Call SetParserError("Uncaught SyntaxError: Unexpected boolean in JSON of " & Char & " at " & index)
            End If
        End If
    End Function

    Private Function parseString(ByRef str, ByRef index)
        Dim Quote, Char, String
        Quote = Mid(str,index,1)
        Do While index <= Len(str)
            index = index + 1
            If index > Len(str) Then
                Call SetParserError("Uncaught SyntaxError: Unexpected end of JSON input, missing '" & Quote & "'")
            End If
            Char = Mid(str,index,1)
            If Char<>Quote Then
                String = String & Char
            Else
                If Mid(str,index-1,1)="\" Then
                    String = String & Char
                Else
                    Exit Do
                End If
            End If
        Loop
        parseString = DecodeUnicode(String)
    End Function

    Private Function parseNumber(ByRef str, ByRef index)
        Dim Char, NumberStr
        Do While index <= Len(str)
            Char = Mid(str,index,1)
            If InStr("+-0123456789.eE", Char) Then
                NumberStr = NumberStr & Char
            Else
                index = index - 1
                Exit Do
            End If
            index = index + 1
        Loop
        Dim RegMatch:Set RegMatch = NumberRegExp.Execute(NumberStr)
        If RegMatch.Count=1 Then
            parseNumber = NumberStr - 0
        Else
            Call SetParserError("Uncaught SyntaxError: Unexpected number in JSON of " & NumberStr & " at " & index)
        End If
    End Function
END Class