'//--Jsonテキストの読込用クラス



'//--
    Private JsonDic As Object
    Private ArrayCountDic As Object
    Private TemplateProp As String
    Private JsonEndIndex As Long
    Private IsLowCaseKey As Boolean
    Private ReadIndex As Long
    Private ReadKeyInfo As String
    Private Const MaxValue As Long = 2147483647

'/--Keyが存在しているか確認する。
Public Function Exists(ByVal Key As String) As Boolean
    If IsNull(JsonDic) Then
        Exists = False
    Else
        Exists = JsonDic.Exists(RewriteKey(Key))
    End If
End Function

'/--Keyが存在しているか確認する。
Public Function Exists2(ParamArray Keys() As Variant) As Boolean
    Exists2 = Exists(BuildKey(Keys()))
End Function

'/--対象のKey配列が配列の場合に最大のIndex番号を返却する。見つからない場合は-1
Public Function ExistsArrayCount(ByVal Key As String) As Long
    If IsNull(JsonDic) Then
        ExistsArrayCount = False
    Else
        Key = RewriteKey(Key)

        If ArrayCountDic.Exists(Key) Then
            ExistsArrayCount = ArrayCountDic.Item(Key)
        Else
            ExistsArrayCount = -1
        End If
    End If
End Function

'/--対象のKey配列が配列の場合に最大のIndex番号を返却する。見つからない場合は-1
Public Function ExistsArrayCount2(ParamArray Keys() As Variant) As Long
    ExistsArrayCount2 = ExistsArrayCount(BuildKey(Keys))
End Function

'/--読み込んでいる#Index番号を取得する。
Public Property Get ActiveReadIndex() As Long
    ActiveReadIndex = ReadIndex
End Property

'/--
Public Property Get GetReWriteTemplateProperty(ByVal ConcatText As String) As String
    GetReWriteTemplateProperty = RewriteKey("") & ConcatText
End Property

'/--各種Readプロパティに結合されるKey情報を登録する。
Public Property Let SetTemplateProperty(ByVal Key As String)
    TemplateProp = Key
End Property

'/--既存のテンプレートKeyに更につなげる形で結合されるKey情報を再登録する。
Public Property Let ConcatNowTemplateProp(ByVal Key As String)
    TemplateProp = TemplateProp & Key
End Property

'/--テンプレートとして指定しているKey情報を開放する。
Public Sub ReleaseTemplateProperty()
    TemplateProp = ""
End Sub

Public Function Loop_AppendTemplate(ParamArray Keys() As Variant) As Boolean
    Dim Key As String
        Key = BuildKey(Keys)
    
    If ReadIndex = -1 Or ReadKeyInfo <> Key Then
        SetTemplateProperty = Key & "#Index"
        ReadKeyInfo = Key
        ReadIndex = 0
    Else
        ReadIndex = ReadIndex + 1
    End If
    
    If ArrayCountDic.Exists(ReadKeyInfo) = False Then GoTo DisableKey
    If ReadIndex > ArrayCountDic.Item(ReadKeyInfo) Then GoTo DisableKey
    
    Loop_AppendTemplate = True
    Exit Function
    
DisableKey:

    ReadIndex = -1
    ReadKeyInfo = ""
    Call ReleaseTemplateProperty
    Loop_AppendTemplate = False
    
End Function
'/--#Indexを[0]に変換したKey情報が見つかった場合はTrueを返す。
'/--更に単発系Readメソッドで#Indexの値を[数値]に変換して取得を行う。
'/--同じKeyを2回以上呼び出している場合、Keyが存在した場合はIndexを加算して変換を行う。

 
'/--値を取得する。
Public Function Read(ByVal Key As String) As Variant
    If Exists(Key) Then
        Read = JsonDic.Item(RewriteKey(Key))
    Else
        Read = Null
    End If
End Function

'/--値を取得する。
Public Function Read2(ParamArray Keys() As Variant) As Variant
    Read2 = Read(BuildKey(Keys()))
End Function

'/--値を取得する。
Public Function Number(ByVal Key As String) As Long
On Error GoTo DisableCast
    Dim Result As Variant
        Result = Read(Key)
        
    If IsNull(Result) Then
DisableCast:
        Number = -99999999
    Else
        Number = CLng(Result)
    End If
End Function

'/--値を取得する。
Public Function Number2(ParamArray Keys() As Variant) As Long
    Number2 = Number(BuildKey(Keys()))
End Function

'/--値を取得する。
Public Function Dbl(ByVal Key As String) As Double
On Error GoTo DisableCast
    Dim Result As Variant
        Result = Read(Key)
    
    If IsNull(Result) Then
DisableCast:
        Dbl = -99999999
    Else
        Dble = CDbl(Result)
    End If
End Function

'/--値を取得する。
Public Function Dbl2(ParamArray Keys() As Variant) As Double
    Dbl2 = Dbl(BuildKey(Keys()))
End Function

'/--値を取得する。
Public Function Text(ByVal Key As String) As String
On Error GoTo DisableCast
    Dim Result As Variant
        Result = Read(Key)
        
    If IsNull(Result) Then
DisableCast:
        Text = ""
    Else
        Text = CStr(Result)
    End If
End Function

'/--値を取得する。
Public Function Text2(ParamArray Keys() As Variant) As String
    Text2 = Text(BuildKey(Keys()))
End Function

'/--値を取得する。
Public Function DateTime(ByVal Key As String) As Date
On Error GoTo DisableCast
    Dim Result As Variant
        Result = Read(Key)
        
    If IsNull(Result) Then
DisableCast:
        DateTime = DateSerial(1900, 1, 1)
    Else
        DateTime = CDate(Result)
    End If
End Function

'/--値を取得する。
Public Function DateTime2(ParamArray Keys() As Variant) As Date
    DateTime2 = DateTime(BuildKey(Keys()))
End Function

'/--値を取得する。
Public Function Bool(ByVal Key As String) As Boolean
On Error GoTo DisableCast
    Dim Result As Variant
        Result = Read(Key)
        
    If IsNull(Result) Then
DisableCast:
        Bool = False
    Else
        Bool = CBool(Result)
    End If
End Function

'/--値を取得する。
Public Function Bool2(ParamArray Keys() As Variant) As Boolean
    Bool2 = Bool(BuildKey(Keys()))
End Function

'/--Keyに紐づくArrayの値を取得する。
'/--取得したいプロパティが『Data->User->ID】の場合、そのままKeyを記載するか、
'/--引数にそれぞれ Data,User,IDと記載してもよい。
Public Function Read_Array(ParamArray Keys() As Variant) As Variant
    Dim Key As String
        Key = BuildKey(Keys)
        
    If Exists(Key & "[0]") Then
        Dim ResultCol As Collection
            Set ResultCol = New Collection
        Dim ArrayIndex As Long
            ArrayIndex = 0
        
        Do While Exists(Key & "[" & ArrayIndex & "]")
            ResultCol.Add Read(RewriteKey(Key & "[" & ArrayIndex & "]"))
            ArrayIndex = ArrayIndex + 1
        Loop
        
        Read_Array = ConvertCollection(ResultCol)
    Else
        Read_Array = Null
    End If
End Function

'/--同階層にある特定のプロパティ値をArrayとして取得する。
Public Function Read_PropsArray(ByVal BaseKey As String, ByVal MainKey As String) As Variant
    If Exists(BaseKey & "[0]->" & MainKey) Then
        Dim ResultCol As Collection
            Set ResultCol = New Collection
        Dim ArrayIndex As Long
            ArrayIndex = 0
        
        Do While Exists(BaseKey & "[" & ArrayIndex & "]->" & MainKey)
            ResultCol.Add Read(RewriteKey(BaseKey & "[" & ArrayIndex & "]->" & MainKey))
            ArrayIndex = ArrayIndex + 1
        Loop
    
        Read_PropsArray = ConvertCollection(ResultCol)
    Else
        Read_PropsArray = Null
    End If
End Function

'/--同階層にある特定のプロパティ値をArrayとしてKeyの数だけ取得する。
'/--取得したい配列部分は[]で記載する。
'/--{A:1,B:[1,2,3,4,5]、C:{CA:[11,22,33,44,55]}} でBの配列を取得したい場合は…B[]とKey入力する。
'/--CAの配列を取得したい場合はC-CA[]
Public Function Read_MultiPropsArray(ParamArray Base_MainKeys() As Variant) As Variant
    Dim Item As Variant, Key As String
    
    Dim Result As Collection
        Set Result = New Collection
        
    For Each Item In Base_MainKeys
        If IsEmpty(Item) Then
            Result.Add Array()
            GoTo NextFor
        End If
        
        Key = CStr(Item)
        
        If InStr(1, Key, "[]") = 0 Then
            Result.Add Read_Array(Key)
        
        Else
            Dim SpKeys As Variant
                SpKeys = Split(Key, "[]")
            
            If Len(SpKeys(0)) = 0 Then
                Result.Add Read_PropsArray("", CStr(SpKeys(1)))
            
            ElseIf Len(SpKeys(1)) = 0 Then
                Result.Add Read_Array(CStr(SpKeys(1)))
            
            Else
                Result.Add Read_PropsArray(CStr(SpKeys(0)), CStr(SpKeys(1)))
            
            End If
        End If
NextFor:
    Next
    
    Dim ResultValues() As Variant
    Dim ColCount As Long
    Dim MaxLen As Long
    
    For Each Item In Result
        If IsNull(Item) Then GoTo ContinueFor
        If MaxLen < UBound(Item) Then MaxLen = UBound(Item)
        ColCount = ColCount + 1
ContinueFor:
    Next
    
    If ColCount = 0 Then
        Read_MultiPropsArray = Null
        Exit Function
    End If
        
    ReDim ResultValues(0 To MaxLen, 0 To ColCount - 1)
    ColCount = 0
    
    Dim i As Long
    For Each Item In Result
        If IsNull(Item) Then GoTo ContinueFor2
        
        For i = 0 To UBound(Item)
            ResultValues(i, ColCount) = Item(i)
        Next
        ColCount = ColCount + 1
ContinueFor2:
    Next
        
    Read_MultiPropsArray = ResultValues
End Function


'/--Collectionを配列に変換する。
Private Function ConvertCollection(ByRef Col As Collection) As Variant
    If Col.Count = 0 Then: Exit Function
    
    Dim Item As Variant, Result As Variant
        ReDim Result(0 To Col.Count - 1)
    
    Dim Index As Long
        Index = 0
    For Each Item In Col
        Result(Index) = Item
        Index = Index + 1
    Next
    
    ConvertCollection = Result
End Function

'/--Keysの情報から１つのキーを作成し返却する。
Private Function BuildKey(ByVal Keys As Variant) As String
    Dim Key As String
    Dim i As Long
    For i = LBound(Keys) To UBound(Keys)
        If Key = "" Then
            Key = CStr(Keys(i))
        Else
            Key = Key & "->" & Keys(i)
        End If
    Next
    
    BuildKey = Key
End Function

'/--Keyの情報をクラス内の指定条件に従い書き換えて返却する。
Private Function RewriteKey(ByVal Key As String) As String
    If IsLowCaseKey Then Key = LCase(Key)
    If TemplateProp <> "" Then Key = TemplateProp & Key
    If ReadIndex >= 0 Then Key = Replace(Key, "#Index", "[" & ReadIndex & "]", Compare:=vbTextCompare)
    
    RewriteKey = Key
End Function
'/--保存したキーを読み込む
Public Sub Debug_Dic()
On Error Resume Next
    
    Debug.Print "--------------------------------"
    
    Dim r As Variant, Result As Variant

    Result = JsonDic.Keys()
    For Each r In Result
        Debug.Print r & " : " & JsonDic.Item(CStr(r))
    Next
    
    Debug.Print "-----------------------------END"
    Exit Sub

    For Each r In Result
        Debug.Print r
    Next
    
    Debug.Print "-----------------------------END"
End Sub

'/--ファイル内のテキストを読込み、JSONを解析する。
'/--LowCaseKeyをTrueにした場合は、Keyの大文字/小文字の区別を行わない。
Public Sub ParseJsonFile(ByVal FilePath As String, Optional ByVal LowCaseKey As Boolean = False)
    Dim File As Object, Stream As Object
        Set File = CreateObject("Scripting.FileSystemObject")
        Set Stream = File.OpenTextFile(FilePath, 1, False)
    
    Dim ParseJsonText As String: ParseJsonText = Stream.ReadAll
    
    Stream.Close
    Set Stream = Nothing
    Set File = Nothing
    
    Call Parse(ParseJsonText, LowCaseKey)
End Sub

'/--指定のJSON文字列を解析します。
'/--LowCaseKeyをTrueにした場合は、Keyの大文字/小文字の区別を行わない。
Public Sub Parse(ByRef ParseJsonText As String, Optional ByVal LowCaseKey As Boolean = False)
    IsLowCaseKey = LowCaseKey
    Set JsonDic = CreateObject("Scripting.Dictionary")
    Set ArrayCountDic = CreateObject("Scripting.Dictionary")
    
    ParseJsonText = Mid$(ParseJsonText, SkipSpace(ParseJsonText, 1))
    JsonEndIndex = Len(ParseJsonText)
        
    Dim CheckTarget As String
        CheckTarget = Mid$(ParseJsonText, 1, 1)
    Dim StartIndex As Long
    
    If CheckTarget = "{" Then
        StartIndex = Parse_Object(Mid$(ParseJsonText, SkipSpace(ParseJsonText, 2)), "")
    
    ElseIf CheckTarget = "[" Then
        StartIndex = Parse_Array(Mid$(ParseJsonText, SkipSpace(ParseJsonText, 2)), "")
    
    Else: Err.Raise 1000, "先端記号が正しくありません。"
    
    End If

    If JsonEndIndex > StartIndex Then
        Err.Raise 1000, "何か処理漏れがありそうですよ…?"
    End If
End Sub


'/--Object形式のデータを処理する。
Private Function Parse_Object(ByRef Text As String, ByVal DicText As String) As Long
    Dim ReadIndex As Long, EndIndex As Long
        ReadIndex = 1
        EndIndex = Len(Text)
    
    Dim CheckTarget As String
        
    Do While ReadIndex < EndIndex
        ReadIndex = SkipSpace(Text, ReadIndex)
        CheckTarget = Mid$(Text, ReadIndex, 1)
        ReadIndex = ReadIndex + 1
        
        If CheckTarget = "{" Then
            ReadIndex = ReadIndex + Parse_Object(Scope(Text, ReadIndex, GetEndSymbolIndex(Text, ReadIndex, CheckTarget)), DicText)
        
        ElseIf CheckTarget = "[" Then
            ReadIndex = ReadIndex + Parse_Array(Scope(Text, ReadIndex, GetEndSymbolIndex(Text, ReadIndex, CheckTarget)), DicText)
        
        ElseIf CheckTarget = "}" Then
            Exit Do
        
        ElseIf CheckTarget = "," Then
            
        Else
            If CheckTarget <> """" Then Err.Raise 1000, "プロパティが存在しません。"
            ReadIndex = SkipSpace(Text, ReadIndex)
            
            Dim SymbolIndex As Long
                SymbolIndex = InStr(ReadIndex + 1, Text, """", vbTextCompare)
                
            '--エスケープ判定
            Do While InStr(1, Mid$(Text, SymbolIndex - 1, 2), "\", vbTextCompare) = 1
                SymbolIndex = InStr(SymbolIndex + 1, Text, """", vbTextCompare)
                If SymbolIndex = 0 Then Err.Raise 1000, "プロパティ終端の記号が正しくありません。"
            Loop
            
            Dim AppendDicText As String
                AppendDicText = Scope(Text, ReadIndex, SymbolIndex)
            
            SymbolIndex = InStr(SymbolIndex, Text, ":", vbTextCompare) + 1
            If SymbolIndex = 1 Then Err.Raise 1000, "プロパティ終端の記号が正しくありません。"
            
            Dim PropScope As String
                PropScope = Mid$(Text, SymbolIndex)
                
            If DicText = "" Then
                ReadIndex = SymbolIndex + Parse_Core(PropScope, AppendDicText)
            Else
                ReadIndex = SymbolIndex + Parse_Core(PropScope, DicText & "->" & AppendDicText)
            End If
            
        End If
    Loop
    
    Parse_Object = EndIndex + 1

End Function


'/--Array形式のデータを処理する。
Private Function Parse_Array(ByRef Text As String, DicText As String) As Long
    Dim ReadIndex As Long, EndIndex As Long
        ReadIndex = 1
        EndIndex = Len(Text)
    
    Dim CheckTarget As String
        
    Dim ArrayIndex As Long
        ArrayIndex = 0
    Dim ArrayDic As String
    
    Do While ReadIndex < EndIndex
        ArrayDic = DicText & "[" & ArrayIndex & "]"
        ReadIndex = SkipSpace(Text, ReadIndex)
        CheckTarget = Mid$(Text, ReadIndex, 1)
        ReadIndex = ReadIndex + 1
        
        If CheckTarget = "{" Then
            ReadIndex = _
                ReadIndex + Parse_Object(Scope(Text, ReadIndex, GetEndSymbolIndex(Text, ReadIndex, CheckTarget)), ArrayDic)
        
        ElseIf CheckTarget = "[" Then
            ReadIndex = _
                ReadIndex + Parse_Array(Scope(Text, ReadIndex, GetEndSymbolIndex(Text, ReadIndex, CheckTarget)), ArrayDic)
        
        ElseIf CheckTarget = "]" Then
            Exit Do
            
        ElseIf CheckTarget = "," Then
            GoTo Continue
        Else
            ReadIndex = ReadIndex + Parse_Core(Mid(Text, ReadIndex - 1), ArrayDic)
        End If
        
        ArrayIndex = ArrayIndex + 1
Continue:
    Loop
    
    ArrayCountDic.Add DicText, ArrayIndex
    Parse_Array = EndIndex + 1
    
End Function

'/--プロパティ値のデータを処理する。
Private Function Parse_Core(ByRef Text As String, ByVal DicText As String) As Long
    If IsLowCaseKey Then DicText = LCase(DicText)
    
    Dim ReadIndex As Long
        ReadIndex = SkipSpace(Text, 1)
        
    Dim CheckTarget As String
        CheckTarget = Mid$(Text, ReadIndex, 1)
               
    Dim AppendIndex As Long
    Select Case CheckTarget
        Case ",": AppendIndex = 1
        Case "{":
            AppendIndex = Parse_Object(Scope(Text, ReadIndex + 1, GetEndSymbolIndex(Text, ReadIndex, CheckTarget)), DicText)
        Case "[":
            AppendIndex = Parse_Array(Scope(Text, ReadIndex + 1, GetEndSymbolIndex(Text, ReadIndex, CheckTarget)), DicText)
        Case "t": JsonDic.Add DicText, True
            AppendIndex = 4
        Case "f": JsonDic.Add DicText, False
            AppendIndex = 5
        Case "n": JsonDic.Add DicText, Null
            AppendIndex = 4
        
        Case Else
            
            Dim SymbolIndex As Long
                 SymbolIndex = GetMinEndSymbolIndex(Text, ReadIndex)
                 
            If SymbolIndex < 1 Or SymbolIndex - ReadIndex < 1 Then Err.Raise 1000, "終端記号が正しくありません。"
            
            Dim Value As String
                Value = Replace(Replace(Mid$(Text, ReadIndex, SymbolIndex - ReadIndex), "\""", "'"), """", "")
            
            If CheckTarget = """" Then GoTo NotNumber
            If IsNumeric(Value) = False Then GoTo NotNumber
            If Len(Value) < 10 Then
                Call AddDic(DicText, CLng(Value))
                'JsonDic.Add DicText, CLng(Value)
            
            ElseIf Len(Value) = 10 Then
                If CLng(Mid$(Value, 1, 9)) > 214748364 Then GoTo NotNumber
                 Call AddDic(DicText, CLng(Value))
                 'JsonDic.Add DicText, CLng(Value)
                 
            End If
            
            GoTo ContinueCore
            
NotNumber:
            Dim CheckUtcValues As Variant
                CheckUtcValues = Split(Value, "T")
            
            If UBound(CheckUtcValues) < 1 Then GoTo NotDate
            If Len(CheckUtcValues(0)) <> 10 Then GoTo NotDate
            If Mid$(CheckUtcValues(0), 5, 1) = "-" And Mid$(CheckUtcValues(0), 8, 1) = "-" Then
                Dim UtcDateValues As Variant
                    UtcDateValues = Split(CStr(CheckUtcValues(0)), "-")
                
                Dim UtcDate As Date
                    UtcDate = DateSerial(CInt(UtcDateValues(0)), CInt(UtcDateValues(1)), CInt(UtcDateValues(2)))
                
                UtcDate = DateAdd("h", CDbl(Mid$(CStr(CheckUtcValues(1)), 1, 2)), UtcDate)
                UtcDate = DateAdd("n", CDbl(Mid$(CStr(CheckUtcValues(1)), 4, 2)), UtcDate)
                UtcDate = DateAdd("s", CDbl(Mid$(CStr(CheckUtcValues(1)), 7, 2)), UtcDate)
                
                Call AddDic(DicText, UtcDate)
                'JsonDic.Add DicText, UtcDate
                GoTo ContinueCore
            End If

NotDate:
            Call AddDic(DicText, Value)
            'JsonDic.Add DicText, Value
ContinueCore:
            ReadIndex = SymbolIndex - ReadIndex
    End Select

    Parse_Core = ReadIndex + AppendIndex
End Function

Private Sub AddDic(ByVal Key As String, Value)
    If JsonDic.Exists(Key) = False Then JsonDic.Add Key, Value
End Sub

'/--記号の終端の次の文字Indexを返却する。
Private Function GetEndSymbolIndex(ByRef Text As String, ByVal Index As Long, ByVal Target As String) As Long
    If Index = 0 Then Err.Raise 1000, "Indexは0以上を指定してください。"
    
    Dim EndTarget As String
    If Target = "{" Then
         EndTarget = "}"
    ElseIf Target = "[" Then
        EndTarget = "]"
    Else
        Err.Raise 1000, "記号が正しくありません。"
    End If
    
    Dim t1 As Long, T2 As Long
        t1 = Index
        T2 = Index
    Do
        t1 = InStr(t1 + 1, Text, Target)
        T2 = InStr(T2 + 1, Text, EndTarget)
    Loop While t1 <> 0 And T2 <> 0 And t1 < T2
    
    If T2 = 0 Then Err.Raise 1000, "終端記号が見つかりません。"
    
    GetEndSymbolIndex = T2 + 1
End Function

'/--終端記号(,}])の内最小のIndexを返却する。
Private Function GetMinEndSymbolIndex(ByRef Text As String, ByVal Index As Long) As Long

    Dim S1 As Long, S2 As Long, S3 As Long
        S1 = InStr(Index, Text, ",", vbTextCompare)
        S2 = InStr(Index, Text, "}", vbTextCompare)
        S3 = InStr(Index, Text, "]", vbTextCompare)
    
    Dim MinIndex As Long
        MinIndex = WorksheetFunction.Min(S1, S2, S3)

    If MinIndex <> 0 Then
        GetMinEndSymbolIndex = MinIndex
    Else
        If S1 = 0 Then S1 = MaxValue
        If S2 = 0 Then S2 = MaxValue
        If S3 = 0 Then S3 = MaxValue
        
        GetMinEndSymbolIndex = WorksheetFunction.Min(S1, S2, S3)
    End If
        
    If GetMinEndSymbolIndex = MaxValue Then Err.Raise 1000, "区切りが正しく終了していません。"
End Function

'/--空白をスキップして次の値を返却する。
Private Function SkipSpace(ByRef Text As String, ByVal Index As Long) As Long
    Dim Target As String
    Target = Mid$(Text, Index, 1)

    Do While JsonEndIndex > Index And (Target = vbCrLf Or Target = vbCr Or Target = vbLf Or Target = vbTab Or Target = vbBack Or Target = vbNewLine Or Target = " " Or Target = "　")
        Index = Index + 1
        Target = Mid$(Text, Index, 1)
    Loop
    
    SkipSpace = Index
End Function

'/--開始～終了までの範囲の文字列を取得する。
Private Function Scope(ByRef Text As String, ByVal StartIndex As Long, ByVal EndIndex As Long) As String
    Scope = Mid$(Text, StartIndex, EndIndex - StartIndex)
End Function

Private Sub Class_Initialize()
    ReadIndex = -1
    ReadKeyInfo = ""
End Sub
