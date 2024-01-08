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
    

    
'/--値を取得する。
Public Function Read(ByVal Key) As Variant
    If JsonDic.Exists(Key) Then
        Read = JsonDic.Item(Key)
    Else
        Read = Null
    End If
End Function

'/--Keyに紐づくArrayの値を取得する。
Public Function Read_Array(ByVal Key) As Variant()
    If JsonDic.Exists(Key & "[0]") Then
        Dim ResultCol As Collection
            Set ResultCol = New Collection
        Dim ReadIndex As Long
            ReadIndex = 0
        
        Do While JsonDic.Exists(Key & "[" & ReadIndex & "]")
            ResultCol.Add JsonDic.Item(Key & "[" & ReadIndex & "]")
            ReadIndex = ReadIndex + 1
        Loop
        
        Read_Array = ConvertCollection(ResultCol)
    Else
        Read_Array = Null
    End If
End Function

'/--同階層にある特定のプロパティ値をArrayとして取得する。
Public Function Read_PropsArray(ByVal BaseKey, ByVal MainKey) As Variant()
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

'/--Object形式のデータを処理する。
Private Function Parse_Object(ByRef Text As String, ByVal Index As Long, ByVal EndIndex As Long, ByVal DicText As String) As Long
    Dim ReadIndex As Long
        ReadIndex = Index
    
    Dim CheckTarget As String
        CheckTarget = Mid(Text, ReadIndex, 1)
    
    Do While ReadIndex <= EndIndex
        ReadIndex = SkipSpace(Text, Index)
        
        If CheckTarget = "{" Then
            ReadIndex = Parse_Object(Text, ReadIndex, GetEndSymbolIndex(Text, ReadIndex + 1, CheckTarget), DicText)
        
        ElseIf CheckTarget = "[" Then
            ReadIndex = Parse_Array(Text, ReadIndex, GetEndSymbolIndex(Text, ReadIndex + 1, CheckTarget), DicText)
        
        Else
            If CheckTarget <> """" Then Err.Raise 1000, "プロパティが存在しません。"
            ReadIndex = SkipSpace(Text, ReadIndex)
            
            Dim SymbolIndex As Long
                SymbolIndex = InStr(ReadIndex + 1, Text, """", vbTextCompare)
                
            '--エスケープ判定
            Do While InStr(SymbolIndex, Text, "\", vbTextCompare) = 1
                SymbolIndex = InStr(SymbolIndex + 1, Text, """", vbTextCompare)
                If SymbolIndex = 0 Then Err.Raise "プロパティ終端の記号が正しくありません。"
            Loop
            
            Dim AppendDicText As String
                AppendDicText = Mid(Text, ReadIndex + 1, SymbolIndex - ReadIndex - 1)
            
            SymbolIndex = InStr(SymbolIndex, Text, ":", vbTextCompare)
            If SymbolIndex = 0 Then Err.Raise "プロパティ終端の記号が正しくありません。"
            
            If DicText = "" Then
                ReadIndex = Parse_Core(Text, SkipSpace(Text, SymbolIndex + 1), AppendDicText)
            Else
                ReadIndex = Parse_Core(Text, SkipSpace(Text, SymbolIndex + 1), DicText & "->" & AppendDicText)
            End If
        End If
        
        If JsonEndIndex >= ReadIndex Then
            Parse_Object = ReadIndex
            Exit Function
        End If
    Loop

End Function


'/--Array形式のデータを処理する。
Private Function Parse_Array(ByRef Text As String, ByVal Index As Long, ByVal EndIndex As Long, DicText As String) As Long

End Function

'/--プロパティ値のデータを処理する。
Private Function Parse_Core(ByRef Text As String, ByVal Index As Long, ByVal DicText As String) As Long
    Dim CheckTarget As String
        CheckTarget = Mid(Text, Index, 1)
    
    Dim ReadIndex As Long
        ReadIndex = Index
        
    Select Case CheckTarget
        Case ",": ReadIndex = Index + 1
        Case "{": ReadIndex = Parse_Object(Text, Index, GetEndSymbolIndex(Text, Index, CheckTarget), DicText)
        Case "[": ReadIndex = Parse_Array(Text, Index, GetEndSymbolIndex(Text, Index, CheckTarget), DicText)
        Case "t": JsonDic.Add DicText, True
        Case "f": JsonDic.Add DicText, False
        Case "n": JsonDic.Add DicText, Null
        
        Case Else
            Dim SymbolIndex As Long
                 SymbolIndex = InStr(ReadIndex, Text, ",", vbTextCompare)
                 ReadIndex = SymbolIndex
                 
            If SymbolIndex < 1 Then Err.Raise 1000, "終端記号が正しくありません。"
            
            If IsNumeric(CheckTarget) Then
                JsonDic.Add DicText, CLng(Mid(Text, ReadIndex, SymbolIndex - ReadIndex))
            Else
            
            End If
            
                
    End Select
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
    
    Dim T1 As Long, T2 As Long
        T1 = Index
        T2 = Index
    Do
        T1 = InStr(T1 + 1, Text, Target)
        T2 = InStr(T2 + 1, Text, EndTarget)
    Loop While T1 <> 0 And T2 <> 0 And T1 < T2
    
    If T2 = 0 Then Err.Raise 1000, "終端記号が見つかりません。"
    
    GetEndSymbolIndex = T2
End Function

'/--空白をスキップして次の値を返却する。
Private Function SkipSpace(ByRef Text As String, ByVal Index As Long) As Long
    Dim Target As String
    Target = Mid(Text, Index, 1)

    Do While JsonEndIndex < Index And (Target = vbCrLf Or Target = vbCr Or Target = vbLf Or Target = vbTab Or Target = vbBack Or Target = vbNewLine Or Target = " " Or Target = "　")
        Index = Index + 1
    Loop
    
    SkipSpace = Index
End Function

Private Sub Class_Initialize()
    Set JsonDic = CreateObject("Scripting.Dictionary")
End Sub
