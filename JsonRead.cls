VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//--Jsonテキストの読込用クラス
'//--文字列を再帰処理で解析します。
'//--デフォルトでは5秒以上経過する場合は強制的に動作を停止させます。
'//--Read系統はKeyを指定し、存在するKeyの場合は値を返します。
'//--Keyの指定方法：①プロパティ名、その直下のプロパティは　①プロパティ名_(アンダースコア)②プロパティ名で指定する。
'//--　　　　　　　 Array系の場合、値内に配列が存在していなければ、1つの配列として保存する。
'//--　　　　　　　 存在している場合は プロパティ名_Array〇〇←初期1からの配列位置
'//--
'//--
'//--例：{"a":1,"b":2,"c":{aa:10,ab:{aaa:100},ac:[10,20,30,40],ad:[10,{ada:1},30,40]}}
'//-- Read("a") -> 1
'//-- Read("c_aa") -> 10
'//-- Read("c_ab_aaa") -> 1000
'//-- Read("c_ac") -> [10,20,30,40]の配列(Variant型)
'//-- Read("c_ad_Array1") -> 10
'//-- Read("c_ad_Array2_ada") -> 1

Private ResultDic As Object
Private ToralTextLength As Long
Private TimeupLimit As Date

Private Sub Class_Initialize()
End Sub

'---値を返します。
Public Property Get Read(ByVal Key As String) As Variant
    If ResultDic.Exists(Key) Then
        Read = ResultDic.Item(Key)
    Else
        Read = -1
    End If
End Property

'---値を返します。
Public Property Get Text(ByVal Key As String) As String
On Error GoTo DisableCast
    If ResultDic.Exists(Key) Then
        Read = CStr(ResultDic.Item(Key))
    Else
DisableCast:
        Read = ""
    End If
End Property

'---値を返します。
Public Property Get Number(ByVal Key As String) As Long
On Error GoTo DisableCast
    If ResultDic.Exists(Key) Then
        Read = CLng(ResultDic.Item(Key))
    Else
DisableCast:
        Read = -1
    End If
End Property

'---値を返します。
Public Property Get Bool(ByVal Key As String) As Boolean
On Error GoTo DisableCast
    If ResultDic.Exists(Key) Then
        Read = CBool(ResultDic.Item(Key))
    Else
DisableCast:
        Read = False
    End If
End Property

'---指定のJSON文字列を解析します。
Public Sub Parse(ByVal ParseJsonText As String)
On Error Resume Next
    Set ResultDic = CreateObject("Scripting.Dictionary")
    
    Call CheckText(ParseJsonText)
    
    ToralTextLength = Len(ParseJsonText)
    Dim StartIndex As Long: StartIndex = 1
    Dim EndIndex As Long: EndIndex = ToralTextLength - 1
    
    '5秒以上の動作は強制停止
    TimeupLimit = DateAdd("s", 500, Now)
    
    Call Parse_Object(ParseJsonText, StartIndex, EndIndex, "")
    
    If EndIndex > StartIndex Then
        Err.Raise 1000, , "何か処理漏れがありそうですよ…?"
    End If
    
    Dim All As Variant: All = ResultDic.Keys()
    Dim AA
    For Each AA In All
        Debug.Print AA & "//:" & ResultDic.Item(AA)
    Next
End Sub

'---{}内のJson文字列に対する処理
Private Sub Parse_Object(ByRef ParseJsonText As String, ByRef StartIndex As Long, ByVal EndIndex As Long, ByVal AddKeyName As String)
    If InStr(StartIndex, ParseJsonText, "{", vbBinaryCompare) <> StartIndex Then
        Err.Raise 1000, , "{で始まっていません…"
        Exit Sub
    Else
        StartIndex = StartIndex + 1
    End If
    
    Dim DicKey As String: If AddKeyName = "" Then DicKey = "" Else DicKey = AddKeyName & "_"
    
    '}終端まで処理を続ける
    Do While StartIndex < EndIndex
        If Mid(ParseJsonText, StartIndex, 1) = "," Then
            StartIndex = StartIndex + 1
            GoTo ContinueFor
        End If
        
        Dim PropIndex As Long: PropIndex = InStr(StartIndex, ParseJsonText, ":")
        Dim PropName As String: PropName = Replace(Mid(ParseJsonText, StartIndex, PropIndex - StartIndex), """", "")
        
        StartIndex = PropIndex + 1
        
        Select Case (Mid(ParseJsonText, StartIndex, 1))
            Case "{":
                Call Parse_Object(ParseJsonText, StartIndex, GetEndSymbolIndex(ParseJsonText, StartIndex, "}"), DicKey & PropName)
                
            Case "[":
                Call Parse_Array(ParseJsonText, StartIndex, GetEndSymbolIndex(ParseJsonText, StartIndex, "]"), DicKey & PropName)
                
            Case Else:
                Call Parse_Core(ParseJsonText, StartIndex, EndIndex, DicKey & PropName)
        End Select
ContinueFor:
    Loop
End Sub

'---[]内のJson文字列に対する処理
Private Sub Parse_Array(ByRef ParseJsonText As String, ByRef StartIndex As Long, ByVal EndIndex As Long, ByVal AddKeyName As String)
    If InStr(StartIndex, ParseJsonText, "[", vbBinaryCompare) <> StartIndex Then
        Err.Raise 1000, , "[で始まっていません…"
        Exit Sub
    Else
        StartIndex = StartIndex + 1
    End If
    
    Dim DicKey As String: If AddKeyName = "" Then DicKey = "" Else DicKey = AddKeyName & "_"
            
    Dim JsonArrayItem As String: JsonArrayItem = Mid(ParseJsonText, StartIndex, EndIndex - StartIndex)
    Dim ResultIndex As Long
    
    '配列系の場合はPropNameにIndex番号を付けて保存する。
    If InStr(1, JsonArrayItem, "{") > 0 Or InStr(1, JsonArrayItem, "[") > 0 Then
        Dim ArrayItemIndex As Long: ArrayItemIndex = 1
                
        Do While StartIndex < EndIndex
            If Mid(ParseJsonText, StartIndex, 1) = "," Then
                StartIndex = StartIndex + 1
                GoTo ContinueFor
            End If
            
            Select Case Mid(JsonArrayItem, StartIndex, 1)
                Case "{":
                    Call Parse_Object(ParseJsonText, StartIndex, GetEndSymbolIndex(ParseJsonText, StartIndex, "}"), AddKeyName & "_Array" & ArrayItemIndex)
                    
                Case "[":
                    Call Parse_Array(ParseJsonText, StartIndex, GetEndSymbolIndex(ParseJsonText, StartIndex, "]"), AddKeyName & "_Array" & ArrayItemIndex)
                    
                Case Else:
                    Call Parse_Core(ParseJsonText, StartIndex, EndIndex, AddKeyName & "_Array" & ArrayItemIndex)
            End Select
            ArrayItemIndex = ArrayItemIndex + 1
ContinueFor:
        Loop
        
    Else
        Dim ArrayItem As Variant
            ArrayItem = Split(Mid(ParseJsonText, StartIndex, EndIndex - StartIndex), ",")
        
        Dim i As Long
        For i = LBound(ArrayItem) To UBound(ArrayItem)
            If IsNumeric(ArrayItem(i)) Then
                ArrayItem(i) = CLng(ArrayItem(i))
            
            ElseIf TypeName(ArrayItem(i)) = "Boolean" Then
                ArrayItem(i) = CBool(ArrayItem(i))
                
            ElseIf InStr(1, CStr(ArrayItem(i)), """") > 0 Then
                ArrayItem(i) = Replace(CLng(ArrayItem(i)), """", "")
            
            End If
        Next
        
        ResultDic.Add Key:=AddKeyName, Item:=ArrayItem
    End If
    
    StartIndex = EndIndex + 2
End Sub

'---プロパティ登録
Private Sub Parse_Core(ByRef ParseJsonText As String, ByRef StartIndex As Long, ByVal EndIndex As Long, ByRef AddDicKey As String)
    If StartIndex < 1 Then
        Err.Raise 1000, , "キーが見つかりませんでした..."
        Exit Sub
    End If
    
    If TimeupLimit < Now Then Err.Raise 1000, , "実行時間が長すぎます…"
    
    Select Case Mid(ParseJsonText, StartIndex, 1)
        Case "{":
            Call Parse_Object(ParseJsonText, StartIndex, GetEndSymbolIndex(ParseJsonText, StartIndex, "}"), AddDicKey)
            
        Case "[":
            Call Parse_Array(ParseJsonText, StartIndex, GetEndSymbolIndex(ParseJsonText, StartIndex, "]"), AddDicKey)
            
        Case Else:
            Dim EndPropIndex As Long: EndPropIndex = InStr(StartIndex, ParseJsonText, ",")
            If EndPropIndex = 0 Then EndPropIndex = StartIndex + 1
            
            Dim JsonItem As String: JsonItem = Replace(Replace(Mid(ParseJsonText, StartIndex, EndPropIndex - StartIndex), "]", ""), "}", "")
            
            If Mid(ParseJsonText, StartIndex, 1) = """" Then
                ResultDic.Add Key:=AddDicKey, Item:=Replace(JsonItem, """", "")
                
            ElseIf InStr(1, JsonItem, "true", vbTextCompare) > 0 Then
                ResultDic.Add Key:=AddDicKey, Item:=True
                
            ElseIf InStr(1, JsonItem, "false", vbTextCompare) > 0 Then
                ResultDic.Add Key:=AddDicKey, Item:=False
                
            ElseIf InStr(1, JsonItem, "null", vbTextCompare) > 0 Then
                ResultDic.Add Key:=AddDicKey, Item:=Null
            
            ElseIf IsNumeric(JsonItem) Then
                ResultDic.Add Key:=AddDicKey, Item:=CLng(JsonItem)
                
            Else
                ResultDic.Add Key:=AddDicKey, Item:=Replace(JsonItem, """", "")
                
            End If
            
            StartIndex = EndPropIndex
    End Select
End Sub

'---最低限の構文チェックを行う
Private Sub CheckText(ByVal Text As String)
    If UBound(Split(Text, "{")) <> UBound(Split(Text, "}")) Then
        Err.Raise 1000, "{}の数が正しくないです…"
    End If
    
    If UBound(Split(Text, "[")) <> UBound(Split(Text, "]")) Then
        Err.Raise 1000, "[]の数が正しくないです…"
    End If
End Sub

'---終端を返す
Private Function GetEndSymbolIndex(ByRef ParseJsonText As String, ByVal StartIndex As Long, ByVal TargetSymbol As String) As Long
    Dim CountSymbol As String: If TargetSymbol = "}" Then CountSymbol = "{" Else CountSymbol = "["
    Dim Count As Long: Count = 0
    
    Dim Count1 As Long, Count2 As Long
        Count2 = StartIndex
        
    Do
        Count1 = InStr(StartIndex, ParseJsonText, CountSymbol)
        Count2 = InStr(Count2 + 1, ParseJsonText, TargetSymbol)
        
    Loop While Count1 <> 0 And Count1 < Count2 And TimeupLimit < Now
    
    GetEndSymbolIndex = Count2
End Function

'---"を取り除く
Private Function ReplaceProp(ByRef Text As String) As String
    ReplaceProp = Replace(Text, """", "", 1, 2)
    If InStr(1, ReplaceProp, """", vbBinaryCompare) <> 0 Then
        Err.Raise 1000, , """がプロパティ内に入っています…"
    End If
End Function
