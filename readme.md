# VBSJSON
一个VBS的JSON数据转换工具类。  
[PS:如果你喜欢的话请给我一个 **Star** ^_^ ]

## 说明
| JSON | VBS |
| --- | --- |
| `null` | `Null` |
| `true` | `True` |
| `false` | `False` |
| `string` | `vbString` |
| `number` | `vbInteger, vbLong, vbSingle, vbDouble` |
| `Array` | `vbArray, vbVariant, vbArray+vbVariant` |
| `Object` | `Dictionary` |

## 使用
```VB
    ' ./example/example.vbs
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim source1, source2
    source1 = FSO.OpenTextFile("source1.json").ReadAll
    source2 = FSO.OpenTextFile("source2.json").ReadAll
    
    Dim JSON
    Set JSON = new VBSJSON

    Dim source1_parse, source2_parse
    Set source1_parse = JSON.parse(source1)
    source2_parse = JSON.parse(source2)

    Dim source1_parse_stringify, source2_parse_stringify
    source1_parse_stringify = JSON.stringify(source1_parse)
    source2_parse_stringify = JSON.stringify(source2_parse)
```

## LICENSE
[MIT License](https://github.com/shihongxins/VBSJSON/blob/main/LICENSE)

## 参考
1. [JSON](https://www.json.org/json-zh.html)
2. [VBJSON](http://www.ediy.co.nz/vbjson-json-parser-library-in-vb6-xidc55680.html)
3. [Demon's VbsJson](http://demon.tw/my-work/vbs-json.html)