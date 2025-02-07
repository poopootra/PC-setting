Attribute VB_Name = "MailDraft"
Sub ShowFileSelectionDialog()
    Dim fileForm As New FileSelectionForm
    Dim selectedFile As String
    Dim folderPath As String
    Dim fullPath As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim rng As Object
    Dim findText As String
    Dim replacementText As String
    Dim outlookApp As Object
    Dim appointmentItem As Object
    Dim mailItem As Object
    Dim pasteContent As String

    folderPath = "C:\Users\mh36264\Documents\04 Templates\Mail\"
    
    ' UserForm を表示
    fileForm.Show
    
    ' 選択されたファイル名を取得
    selectedFile = fileForm.GetSelectedFileName
    
    ' フルパスを生成
    If selectedFile <> "" Then
        fullPath = folderPath & selectedFile
        
        ' Wordを非表示で開く
        Set wordApp = CreateObject("Word.Application")
        wordApp.Visible = False ' Wordを非表示で開く
        Set wordDoc = wordApp.Documents.Open(fullPath)
        
        ' Word文書内の{{}}で囲まれたテキストを検索して置換
        Do
            ' {{}}で囲まれたテキストを検索
            Set rng = wordDoc.Content
            rng.Find.ClearFormatting
            rng.Find.text = "\{\{*\}\}"
            rng.Find.MatchWildcards = True ' ワイルドカード検索を有効にする
            
            If rng.Find.Execute Then
                ' 見つかったテキストを取得
                findText = rng.text
                ' ユーザーに置換後のテキストを尋ねる
                replacementText = InputBox("置換するテキストを入力してください:" & vbCrLf & "見つかったテキスト: " & findText, "テキストの置換")
                
                If replacementText <> "" Then
                    ' テキストを置換
                    rng.text = replacementText
                Else
                    MsgBox "置換後のテキストが入力されませんでした。"
                End If
            Else
                Exit Do ' もう{{}}で囲まれたテキストが見つからなかった場合は終了
            End If
        Loop
        
        ' 変更内容をコピーしてクリップボードに保存
        wordDoc.Content.Copy ' コピーすることでクリップボードに内容を保存
        
        ' Outlookを操作するためのオブジェクトを作成
        Set outlookApp = CreateObject("Outlook.Application")
        
        ' 現在アクティブなメールアイテムまたはアポイントメントアイテムを取得
        On Error Resume Next
        Set mailItem = outlookApp.ActiveInspector.CurrentItem
        Set appointmentItem = outlookApp.ActiveInspector.CurrentItem
        On Error GoTo 0
        
        ' メールアイテムがアクティブな場合
        If Not mailItem Is Nothing Then
            ' クリップボードの内容を貼り付け（書式を保持）
            mailItem.HTMLBody = mailItem.HTMLBody & "<br>" ' HTML本文に改行を追加
            mailItem.GetInspector.CommandBars.ExecuteMso "Paste" ' 貼り付け
            MsgBox "メールにテキストを貼り付けました。"
        ' アポイントメントアイテムがアクティブな場合
        ElseIf Not appointmentItem Is Nothing Then
            ' クリップボードの内容を貼り付け（書式を保持）
            appointmentItem.Body = appointmentItem.Body & vbCrLf ' 本文に改行を追加
            appointmentItem.GetInspector.CommandBars.ExecuteMso "Paste" ' 貼り付け
            MsgBox "アポイントメントにテキストを貼り付けました。"
        Else
            MsgBox "アクティブなメールまたはアポイントメントが見つかりません。"
        End If
        
        ' Wordを保存せずに終了
        wordDoc.Close False
        wordApp.Quit
    Else
        MsgBox "ファイルが選択されませんでした。"
    End If
End Sub

