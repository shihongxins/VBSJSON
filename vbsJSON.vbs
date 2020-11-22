' 
' Author: shihongxins
' Date: 2020-11-21
' GitHub: https://github.com/shihongxins
' LICENSE: MIT License https://github.com/shihongxins/vbsJSON/blob/main/LICENSE
' 

'严格模式：类似与 js 中的 "use strict"
Option Explicit
Class vbsJSON
    '空白符,数字格式正则
    Private WhiteSpace, NumberRegExp
    '错误信息
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
                'BackSpace 退格
                Case 8      uChr = "\b"
                'Tab 缩进
                Case 9      uChr = "\t"
                'Linefeed 换行
                Case 10     uChr = "\n"
                'Formfeed 分页
                Case 12     uChr = "\f"
                'Carriage Return 回车
                Case 13     uChr = "\r"
                'Quotation mark 双引号
                Case 34     uChr = "\"""
                'Single quotation mark 单引号
                Case 39     uChr = "\'"
                'Reverse solidus 反斜杠转义符
                Case 92     uChr = "\\"
                Case Else
                    If uChrCode<32 Or uChrCode>127 Then
                        'non-ascii 非ASCII码
                        uChr = "\u" & Right("0000" & Hex(uChrCode), 4)
                    Else
                        '是ASCII码
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
                    'BackSpace 退格
                    Case "b"    uChr = ChrW(8)
                    'Tab 缩进
                    Case "t"    uChr = ChrW(9)
                    'Linefeed 换行
                    Case "n"    uChr = ChrW(10)
                    'Formfeed 分页
                    Case "f"    uChr = ChrW(12)
                    'Carriage Return 回车
                    Case "r"    uChr = ChrW(13)
                    'Quotation mark 双引号
                    Case """"   uChr = ChrW(34)
                    'Single quotation mark 单引号
                    Case "'"    uChr = ChrW(39)
                    'Solidus 斜杠
                    Case "/"    uChr = ChrW(47)
                    'Reverse solidus 反斜杠转义符
                    Case "\"    uChr = ChrW(92)
                    Case "u"
                        ' Unicode 完整（包括最后的）
                        If i+4 <=Len(str) Then
                            uChr = ChrW("&H" & Mid(str,i + 1,4))
                            i = i + 4
                        Else
                            '补回前面的 '\'
                            uChr = "\" & uChr
                        End If
                End Select
            End If
            uStr = uStr & uChr
        Next
        DecodeUnicode = uStr
    End Function

    'Skip to the next non-blank character
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
            'default: 都没有匹配到，默认执行
            Case Else
                '先转为字符串类型，在转为 unicode
                buf.Add buf.Count, """" & EncodeUnicode(CStr(obj)) & """"
        End Select
        '将执行完后的结果集对象，取其每项的值为数组，然后通过JOIN方法转换为字符串，并返回
        stringify = Join(buf.Items, "")
    End Function

    Public Function parse(ByRef str)
        Call ClearParserError
        '先是类型判断
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
                '字符串格式，退出类型判断，开始解析
            Case Else
                '其他类型报错
                Call SetParserError("Uncaught SyntaxError: Invalid JSON input,Unknown Type")
        End Select

        str = CStr(str)
        '空字符串报错
        If Trim(str)="" Then
            Call SetParserError("Uncaught SyntaxError: Unexpected end of JSON input,Empty String")
        End If
        '开始解析，按不同大类型分类解析，并返回结果
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
        'any character except whitespace after select parse done,will throw error
        'like "[1,2,3]  e  " ,will throw error at 'e'
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
                    'unknown base type
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

        ' Char="{" start parse Key
        Dim obj:Set obj = CreateObject("Scripting.Dictionary")
        KeyEndFlag = True
        Do
            index = index + 1
            If index > Len(str) Then
                Call SetParserError("Uncaught SyntaxError: Unexpected end of JSON input,Missing '}'")
            End If
            If KeyEndFlag=True Then
                'skip whitespace before each key
                index = skipWhiteSpace(str,index)
                Char = Mid(str,index,1)
                'start read property key
                If Char="""" Or Char="'" Then
                    Quote = Char
                    KeyEndFlag = False
                'continue read next property
                ElseIf Char="," Then
                    'throw error like '{,}' and '{,"a":1}' and '{"a":1,,"c":3}'
                    If VarType(Value)=vbEmpty Then
                        Call SetParserError("Uncaught SyntaxError: Unexpected token " & Char & " in JSON at position " & CStr(index))
                    'continue read next property
                    Else
                        'clear last Key and Value
                        Key = ""
                        Value = Empty
                    End If
                'object property done
                ElseIf Char="}" Then
                    'skip "}"
                    'index = index + 1
                    Exit DO
                Else
                    Call SetParserError("Uncaught SyntaxError: Unexpected token " & Char & " in JSON at position " & CStr(index))
                End If
            Else
                Char = Mid(str,index,1)
                If Char<>Quote Then
                    Key = Key + Char
                'stop read property key
                ElseIf Char=Quote Then
                    KeyEndFlag = True
                    index = index + 1
                    'skip whitespace before each ":"
                    index = skipWhiteSpace(str,index)
                    Char = Mid(str,index,1)
                    If Char=":" Then
                        index = index + 1
                        'skip whitespace after each ":"
                        index = skipWhiteSpace(str,index)
                        '判断 Value 类型
                        Dim FirstChar
                        FirstChar = Mid(str,index,1)
                        If FirstChar="{" Then
                            Set Value = parseValue(str,index)
                        Else
                            Value = parseValue(str,index)
                        End If
                        'read Key and Value done,Add to dictionary
                        Key = DecodeUnicode(Key)
                        If obj.Exists(Key) Then
                            '有可能Key 重复且 Value 是对象，不能简单的用赋值，应该先移除再添加
                            'obj.Item(Key) = Value
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

        ' Char="[" start parse
        Dim obj:Set obj = CreateObject("Scripting.Dictionary")
        'skip "["
        index = index + 1
        Do
            'clear all whitespace before every ',' delimiter
            If VarType(Value)=vbEmpty Then
                index = skipWhiteSpace(str,index)
            End If
            If index > Len(str) Then
                Call SetParserError("Uncaught SyntaxError: Unexpected end of JSON input,Missing ']'")
            End If
            Char = Mid(str,index,1)
            If Char="," Then
                'throw error like "[,]" and "[,1]" and "[1,,3]"
                If VarType(Value)=vbEmpty Then
                    Call SetParserError("Uncaught SyntaxError: Unexpected token " & Char & " in JSON at position " & CStr(index))
                'continue read next property
                Else
                    'clear last Value
                    Value = Empty
                End If
            ElseIf Char="]" Then
                'skip "]"
                'index = index + 1
                Exit Do
            Else
                '判断 Value 类型
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

    'parse object / array / null / true / false / string / number
    Private Function parseValue(ByRef str,ByRef index)
        Dim Char
        'maybe doesn't need next line
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
                    'unknown type
                    Call SetParserError("Uncaught SyntaxError: Unexpected token " & Char & " in JSON at position " & CStr(index))
                End If
        End Select
    End Function

    'parse null
    Private Function parseNull(ByRef str, ByRef index)
        If index+4 <= Len(str) And Mid(str,index,4) = "null" Then
            parseNull = Null
            'reset index to last char 'l'
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

    'parse true / false
    Private Function parseBoolean(ByRef str, ByRef index)
        If index+5 <= Len(str) And Mid(str,index,5) = "false" Then
            parseBoolean = False
            'reset index to last char 'e'
            index = index + 4
        ElseIf index+4 <= Len(str) And Mid(str,index,4) = "true" Then
            parseBoolean = True
            'reset index to last char 'e'
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

    'parse string
    Private Function parseString(ByRef str, ByRef index)
        Dim Quote, Char, String
        'Start match, get the opening(left) quotation mark
        Quote = Mid(str,index,1)
        Do While index <= Len(str)
            index = index + 1
            'Until the last character, the match is not over yet
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
                    'End match, skip the closing(right) quotation mark
                    'fixed:not need, beacause the index will change in Do...Loop at parseObject/parseArray
                    'index = index + 1
                    Exit Do
                End If
            End If
        Loop
        parseString = DecodeUnicode(String)
    End Function

    'parse number
    Private Function parseNumber(ByRef str, ByRef index)
        Dim Char, NumberStr
        Do While index <= Len(str)
            Char = Mid(str,index,1)
            If InStr("+-0123456789.eE", Char) Then
                NumberStr = NumberStr & Char
            'untill last char,parseString() still run
            'ElseIf index = Len(str) Then
            '    Call SetParserError("Uncaught SyntaxError: Unexpected end of JSON input")
            Else
                'reset index to last number char
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