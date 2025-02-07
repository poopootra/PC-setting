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
    
    ' �V�X�e�����ϐ�����API�L�[���擾
    apiKey = Environ("OPENAI_API_KEY")
        
    ' API���N�G�X�g��URL�iChatGPT API�̃G���h�|�C���g�j
    url = "https://api.openai.com/v1/chat/completions"
    
    ' �V�X�e���v�����v�g���܂߂�JSON���N�G�X�g�쐬
    jsonRequest = "{""model"":""gpt-4o-mini"",""messages"":[{""role"":""system"",""content"":""���Ȃ��͗D�G�ȃr�W�l�X�}���Ńr�W�l�X�V�[���Ń��[���������Ă��܂��B�����烆�[�U�[���ȑf�����ꂽ���[���̗v�_�݂̂���͂���̂ŁA������r�W�l�X���[���Ƃ��č쐬���Ă��������B�o�͂̓��[���̖{���݂̂Ƃ��Ă��������B�����╶���̏����͕s�v�ł��B�K�X���s�����Ă��������B""},{""role"":""user"",""content"":""" & userInput & """}],""temperature"":0}"
    
    
    ' API�L�[���ݒ肳��Ă��Ȃ��ꍇ�A�G���[���b�Z�[�W��\��
    If apiKey = "" Then
        MsgBox "API Key is not set in the system environment variables."
        Exit Sub
    End If
    
    ' �����s���͂��擾���邽�߂�InputBox�����Ɏg��
    userInput = InputBox("Enter key points for the business email:", "Business Email Input", "Type your key points here...")
    
    ' ��̏ꍇ�͏����𒆒f
    If userInput = "" Then
        MsgBox "No input provided!"
        Exit Sub
    End If
    
    ' HTTP���N�G�X�g�I�u�W�F�N�g�̍쐬
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    xmlhttp.Open "POST", url, False
    
    ' �w�b�_�[�ݒ�
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.setRequestHeader "Authorization", "Bearer " & apiKey
    
    ' ���N�G�X�g���M
    xmlhttp.Send jsonRequest
    
    ' ���X�|���X���擾
    jsonResponse = xmlhttp.responseText
    
    ' �f�o�b�O�̂��߂�JSON���X�|���X���C�~�f�B�G�C�g�E�B���h�E�ɏo��
    Debug.Print "JSON Response: " & jsonResponse
    
    ' ���X�|���X����͂��ĕԓ����������o��
    On Error Resume Next
    Dim json As Object
    Set json = JsonConverter.ParseJson(jsonResponse)
    
    ' JSON���K�؂Ƀp�[�X����Ă��邩�m�F
    If json Is Nothing Then
        MsgBox "Error parsing JSON response. Raw response: " & jsonResponse
        Exit Sub
    End If
    
    ' ChatGPT�̉������e���擾
    On Error GoTo 0
    response = json("choices")(1)("message")("content")
    
    ' �G���[�`�F�b�N�ƕ\��
    If response <> "" Then
        ' �A�N�e�B�u�ȃ��[���A�C�e�����擾
        Set mailItem = Application.ActiveInspector.CurrentItem
        
        ' ���[�����J���Ă��邩�m�F
        If Not mailItem Is Nothing Then
            Dim objDoc
            Set objDoc = mailItem.GetInspector().WordEditor
            
            
            ' �����̖{���̐擪�Ƀ��X�|���X��}��
            With objDoc.Application
                .Selection.TypeText response
            End With
        Else
            MsgBox "No active mail item found."
        End If
        
    Else
        ' �p�[�X���ꂽ���A���҂����������Ȃ������ꍇ�A���X�|���X�����̂܂ܕ\��
        MsgBox "Unexpected response format. Full JSON response: " & vbCrLf & jsonResponse
    End If
End Sub

