Option Explicit
'HttpRequest送信、Response受取用
'''
    '//HTTPリクエスト用のオブジェクト
    Private Res As Object
'''

'リクエスト用のオブジェクトを生成する。
Private Sub Class_Initialize()
    Set Res = CreateObject("MSXML2.XMLHTTP")
End Sub

''GETリクエスト
''URL:リクエストを送るURL
''ContentType:渡す型:application/json等が初期値
Public Sub httpGet(ByVal URL As String, _
    Optional ByVal ContentType As String = "application/json")
    With Res
        .Open "GET", URL, False
        .setRequestheader "Content-Type", ContentType
        .Send
    End With
End Sub

''POSTリクエスト
''URL:リクエストを送るURL
''Data:Postするデータ
''ContentType:渡す型:application/json等が初期値
Public Sub httpPost( _
    ByVal URL As String, _
    ByVal Data As Variant, _
    Optional ByVal ContentType As String = "application/json")
    
    With Res
        .Open "POST", URL, False
        .setRequestheader "Content-Type", ContentType
        .Send Data
    End With
End Sub

''要求返却
Public Function GetResponse() As String
    Dim Result As String: Result = Res.ResponseText
    Dim Reg As Object
        Set Reg = CreateObject("VBScript.RegExp")
    
    With Reg
        .Global = True
        .IgnoreCase = True
        .Pattern = "\\u([a-fA-F0-9]{4})"
    End With
    
    Dim ResultReg As Object 'マッチオブジェクト
        Set ResultReg = Reg.Execute(Result)
    
    Dim Matchs As Variant, SubMatchs As Variant
    For Each Matchs In ResultReg
        For Each SubMatchs In Matchs.SubMatches
            Result = Replace(Result, "\u" & SubMatchs, ChrW("&h" & SubMatchs))
        Next
    Next
    
    Set Reg = Nothing
    
    GetResponse = Result
End Function

