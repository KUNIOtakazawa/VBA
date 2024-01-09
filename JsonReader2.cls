VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonReader2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


'//--
    Private JsonDic As Object
    Private JsonEndIndex As Long
    Private Const MaxValue As Long = 2147483647
    
 
'/--値を取得する。
Public Function Read(ByVal key) As Variant
    If JsonDic.Exists(key) Then
        Read = JsonDic.Item(key)
    Else
        Read = Null
    End If
End Function

'/--Keyに紐づくArrayの値を取得する。
Public Function Read_Array(ByVal key) As Variant()
    If JsonDic.Exists(key & "[0]") Then
        Dim ResultCol As Collection
            Set ResultCol = New Collection
        Dim ReadIndex As Long
            ReadIndex = 0
        
        Do While JsonDic.Exists(key & "[" & ReadIndex & "]")
            ResultCol.Add JsonDic.Item(key & "[" & ReadIndex & "]")
            ReadIndex = ReadIndex + 1
        Loop
        
        Read_Array = ConvertCollection(ResultCol)
    Else
        Read_Array = Null
    End If
End Function

'/--同階層にある特定のプロパティ値をArrayとして取得する。
Public Function Read_PropsArray(ByVal BaseKey, ByVal MainKey) As Variant
    If JsonDic.Exists(BaseKey & "[0]->" & MainKey) Then
        Dim ResultCol As Collection
            Set ResultCol = New Collection
        Dim ReadIndex As Long
            ReadIndex = 0
        
        Do While JsonDic.Exists(BaseKey & "[" & ReadIndex & "]->" & MainKey)
            ResultCol.Add JsonDic.Item(BaseKey & "[" & ReadIndex & "]->" & MainKey)
            ReadIndex = ReadIndex + 1
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
Public Function Read_MultiPropsArray(ParamArray Base_MainKeys()) As Variant

End Function


'/--Collectionを配列に変換する。
Private Function ConvertCollection(ByRef Col As Collection) As Variant()
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

'---ファイル内のテキストを読込み、JSONを解析する。
Public Sub ParseJsonFile(ByVal FilePath As String)
    Dim File As Object, Stream As Object
        Set File = CreateObject("Scripting.FileSystemObject")
        Set Stream = File.OpenTextFile(FilePath, 1, False)
    
    Dim ParseJsonText As String: ParseJsonText = Stream.ReadAll
    
    Stream.Close
    Set Stream = Nothing
    Set File = Nothing
    
    Call Parse(ParseJsonText)
End Sub

'/--指定のJSON文字列を解析します。
Public Sub Parse(ByRef ParseJsonText As String)
    Set JsonDic = CreateObject("Scripting.Dictionary")
    
    
    Dim Index As Long
        Index = 1
        JsonEndIndex = Len(ParseJsonText) - 1
    
    Index = SkipSpace(ParseJsonText, Index)
    
    '5秒以上の動作は強制停止
    TimeupLimit = DateAdd("s", 5000, Now)
        
    Dim CheckTarget As String
        CheckTarget = Mid(ParseJsonText, Index, 1)
        Index = Index + 1
        
    If CheckTarget = "{" Then
        StartIndex = Parse_Object(ParseJsonText, Index, JsonEndIndex, "")
    
    ElseIf CheckTarget = "[" Then
        StartIndex = Parse_Array(ParseJsonText, Index, JsonEndIndex, "")
    
    Else: Err.Raise 1000, "先端記号が正しくありません。"
    
    End If

    If JsonEndIndex > StartIndex Then
        Err.Raise 1000, , "何か処理漏れがありそうですよ…?"
    End If
End Sub


'/--Object形式のデータを処理する。
Private Function Parse_Object(ByRef Text As String, ByVal Index As Long, ByVal EndIndex As Long, ByVal DicText As String) As Long
    Dim ReadIndex As Long
        ReadIndex = Index
    
    Dim CheckTarget As String
        
    Do While ReadIndex <= EndIndex
        ReadIndex = SkipSpace(Text, ReadIndex)
        CheckTarget = Mid(Text, ReadIndex, 1)
        ReadIndex = ReadIndex + 1
        
        If CheckTarget = "{" Then
            ReadIndex = Parse_Object(Text, ReadIndex, GetEndSymbolIndex(Text, ReadIndex + 1, CheckTarget), DicText)
        
        ElseIf CheckTarget = "[" Then
            ReadIndex = Parse_Array(Text, ReadIndex, GetEndSymbolIndex(Text, ReadIndex + 1, CheckTarget), DicText)
        
        ElseIf CheckTarget = "}" Then
            Parse_Object = EndIndex + 1
            Exit Function
        
        ElseIf CheckTarget = "," Then
            
        Else
            If CheckTarget <> """" Then Err.Raise 1000, "プロパティが存在しません。"
            ReadIndex = SkipSpace(Text, ReadIndex)
            
            Dim SymbolIndex As Long
                SymbolIndex = InStr(ReadIndex + 1, Text, """", vbTextCompare)
                
            '--エスケープ判定
            Do While InStr(SymbolIndex, Text, "\", vbTextCompare) = 1
                SymbolIndex = InStr(SymbolIndex + 1, Text, """", vbTextCompare)
                If SymbolIndex = 0 Then Err.Raise 1000, "プロパティ終端の記号が正しくありません。"
            Loop
            
            Dim AppendDicText As String
                AppendDicText = Mid(Text, ReadIndex, SymbolIndex - ReadIndex)
            
            SymbolIndex = InStr(SymbolIndex, Text, ":", vbTextCompare)
            If SymbolIndex = 0 Then Err.Raise 1000, "プロパティ終端の記号が正しくありません。"
            
            If DicText = "" Then
                ReadIndex = Parse_Core(Text, SkipSpace(Text, SymbolIndex + 1), AppendDicText)
            Else
                ReadIndex = Parse_Core(Text, SkipSpace(Text, SymbolIndex + 1), DicText & "->" & AppendDicText)
            End If
        End If
        
        If JsonEndIndex <= ReadIndex Then
            Parse_Object = ReadIndex
            Exit Function
        End If
    Loop
    
    Parse_Object = ReadIndex

End Function


'/--Array形式のデータを処理する。
Private Function Parse_Array(ByRef Text As String, ByVal Index As Long, ByVal EndIndex As Long, DicText As String) As Long
    Dim ReadIndex As Long
        ReadIndex = Index
    
    Dim CheckTarget As String
        
    Dim ArrayIndex As Long
        ArrayIndex = 0
    Dim ArrayDic As String
    
    Do While ReadIndex <= EndIndex
        ArrayDic = DicText & "[" & ArrayIndex & "]"
        ReadIndex = SkipSpace(Text, ReadIndex)
        CheckTarget = Mid(Text, ReadIndex, 1)
        ReadIndex = ReadIndex + 1
        
        If CheckTarget = "{" Then
            ReadIndex = Parse_Object(Text, ReadIndex, GetEndSymbolIndex(Text, ReadIndex, CheckTarget), ArrayDic)
        
        ElseIf CheckTarget = "[" Then
            ReadIndex = Parse_Array(Text, ReadIndex, GetEndSymbolIndex(Text, ReadIndex, CheckTarget), ArrayDic)
        
        ElseIf CheckTarget = "]" Then
            Parse_Array = EndIndex + 1
            Exit Function
            
        ElseIf CheckTarget = "," Then
            GoTo Continue
        Else
            ReadIndex = Parse_Core(Text, ReadIndex - 1, ArrayDic)
        End If
        
        If JsonEndIndex <= ReadIndex Then
            Parse_Array = ReadIndex
            Exit Function
        End If
        
        ArrayIndex = ArrayIndex + 1
Continue:
    Loop
    
    Parse_Array = ReadIndex
    
End Function

'/--プロパティ値のデータを処理する。
Private Function Parse_Core(ByRef Text As String, ByVal Index As Long, ByVal DicText As String) As Long
    If JsonEndIndex <= Index Then
        Parse_Core = JsonEndIndex
        Exit Function
    End If
    
    Dim CheckTarget As String
        CheckTarget = Mid(Text, Index, 1)
    
    Dim ReadIndex As Long
        ReadIndex = Index
                
    Select Case CheckTarget
        Case ",": ReadIndex = ReadIndex + 1
        Case "{": ReadIndex = Parse_Object(Text, Index + 1, GetEndSymbolIndex(Text, Index, CheckTarget), DicText)
        Case "[": ReadIndex = Parse_Array(Text, Index + 1, GetEndSymbolIndex(Text, Index, CheckTarget), DicText)
        Case "t": JsonDic.Add DicText, True
            ReadIndex = ReadIndex + 4
        Case "f": JsonDic.Add DicText, False
            ReadIndex = ReadIndex + 5
        Case "n": JsonDic.Add DicText, Null
            ReadIndex = ReadIndex + 4
        
        Case Else
            Dim SymbolIndex As Long
                 SymbolIndex = GetMinEndSymbolIndex(Text, ReadIndex)
                 
            If SymbolIndex < 1 Or SymbolIndex - Index < 1 Then Err.Raise 1000, "終端記号が正しくありません。"
            
            Dim value As String
                value = Replace(Replace(Mid(Text, Index, SymbolIndex - ReadIndex), "\""", "'"), """", "")
            
            If IsNumeric(value) = False Then GoTo NotNumber
            If Len(value) < 10 Then
                JsonDic.Add DicText, CLng(value)
            
            ElseIf Len(value) = 10 Then
                If CLng(Mid(value, 1, 9)) > 214748364 Then GoTo NotNumber
                 JsonDic.Add DicText, CLng(value)
                 
            End If
            
            GoTo ContinueCore
            
NotNumber:
            Dim CheckUtcValues As Variant
                CheckUtcValues = Split(value, "T")
            
            If UBound(CheckUtcValues) < 1 Then GoTo NotDate
            If Len(CheckUtcValues(0)) <> 10 Then GoTo NotDate
            If Mid(CheckUtcValues(0), 5, 1) = "-" And Mid(CheckUtcValues(0), 8, 1) = "-" Then
                Dim UtcDateValues As Variant
                    UtcDateValues = Split(CStr(CheckUtcValues(0)), "-")
                
                Dim UtcDate As Date
                    UtcDate = DateSerial(CInt(UtcDateValues(0)), CInt(UtcDateValues(1)), CInt(UtcDateValues(2)))
                
                UtcDate = DateAdd("h", CDbl(Mid(CStr(CheckUtcValues(1)), 1, 2)), UtcDate)
                UtcDate = DateAdd("n", CDbl(Mid(CStr(CheckUtcValues(1)), 4, 2)), UtcDate)
                UtcDate = DateAdd("s", CDbl(Mid(CStr(CheckUtcValues(1)), 7, 2)), UtcDate)
                
                JsonDic.Add DicText, UtcDate
                GoTo ContinueCore
            End If

NotDate:
    JsonDic.Add DicText, value
ContinueCore:
            
            ReadIndex = SymbolIndex + 1
    End Select
    
    Parse_Core = SkipSpace(Text, ReadIndex)
End Function

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
    
    Dim t1 As Long, t2 As Long
        t1 = Index
        t2 = Index
    Do
        t1 = InStr(t1 + 1, Text, Target)
        t2 = InStr(t2 + 1, Text, EndTarget)
    Loop While t1 <> 0 And t2 <> 0 And t1 < t2
    
    If t2 = 0 Then Err.Raise 1000, "終端記号が見つかりません。"
    
    GetEndSymbolIndex = t2
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
    Target = Mid(Text, Index, 1)

    Do While JsonEndIndex > Index And (Target = vbCrLf Or Target = vbCr Or Target = vbLf Or Target = vbTab Or Target = vbBack Or Target = vbNewLine Or Target = " " Or Target = "　")
        Index = Index + 1
        Target = Mid(Text, Index, 1)
    Loop
    
    SkipSpace = Index
End Function

'/--開始～終了までの範囲の文字列を取得する。
Private Function Scope(ByRef Text As String, ByVal StartIndex As Long, ByVal EndIndex As Long) As String
    Scope = Mid$(Text, StartIndex, EndIndex - 1)
End Function
