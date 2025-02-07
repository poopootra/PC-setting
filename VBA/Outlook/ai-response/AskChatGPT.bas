Attribute VB_Name = "AskChatGPT"
Sub ChatGPTBusinessEmailResponse()
    Dim userInput As String
    Dim response As String
    Dim apiKey As String
    Dim url As String
    Dim jsonRequest As String
    Dim xmlhttp As Object
    Dim jsonResponse As String
    Dim mailItem As Object
    
    ' システム環境変数からAPIキーを取得
    apiKey = Environ("OPENAI_API_KEY")
        
    ' APIリクエストのURL（ChatGPT APIのエンドポイント）
    url = "https://api.openai.com/v1/chat/completions"
    
    ' システムプロンプトを含めたJSONリクエスト作成
    jsonRequest = "{""model"":""gpt-4o-mini"",""messages"":[{""role"":""system"",""content"":""あなたは優秀なビジネスマンでビジネスシーンでメールを書いています。今からユーザーが簡素化されたメールの要点のみを入力するので、それをビジネスメールとして作成してください。出力はメールの本文のみとしてください。件名や文末の署名は不要です。適宜改行を入れてください。""},{""role"":""user"",""content"":""" & userInput & """}],""temperature"":0}"
    
    
    ' APIキーが設定されていない場合、エラーメッセージを表示
    If apiKey = "" Then
        MsgBox "API Key is not set in the system environment variables."
        Exit Sub
    End If
    
    ' 複数行入力を取得するためにInputBoxを代わりに使う
    userInput = InputBox("Enter key points for the business email:", "Business Email Input", "Type your key points here...")
    
    ' 空の場合は処理を中断
    If userInput = "" Then
        MsgBox "No input provided!"
        Exit Sub
    End If
    
    ' HTTPリクエストオブジェクトの作成
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "POST", url, False
    
    ' ヘッダー設定
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.setRequestHeader "Authorization", "Bearer " & apiKey
    
    ' リクエスト送信
    xmlhttp.Send jsonRequest
    
    ' レスポンスを取得
    jsonResponse = xmlhttp.responseText
    
    ' デバッグのためにJSONレスポンスをイミディエイトウィンドウに出力
    Debug.Print "JSON Response: " & jsonResponse
    
    ' レスポンスを解析して返答部分を取り出す
    On Error Resume Next
    Dim json As Object
    Set json = JsonConverter.ParseJson(jsonResponse)
    
    ' JSONが適切にパースされているか確認
    If json Is Nothing Then
        MsgBox "Error parsing JSON response. Raw response: " & jsonResponse
        Exit Sub
    End If
    
    ' ChatGPTの応答内容を取得
    On Error GoTo 0
    response = json("choices")(1)("message")("content")
    
    ' エラーチェックと表示
    If response <> "" Then
        ' アクティブなメールアイテムを取得
        Set mailItem = Application.ActiveInspector.CurrentItem
        
        ' メールが開いているか確認
        If Not mailItem Is Nothing Then
            Dim objDoc
            Set objDoc = mailItem.GetInspector().WordEditor
            
            
            ' 既存の本文の先頭にレスポンスを挿入
            With objDoc.Application
                .Selection.TypeText response
            End With
        Else
            MsgBox "No active mail item found."
        End If
        
    Else
        ' パースされたが、期待した応答がなかった場合、レスポンスをそのまま表示
        MsgBox "Unexpected response format. Full JSON response: " & vbCrLf & jsonResponse
    End If
End Sub

