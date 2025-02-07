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
    
    ' UserForm ��\��
    fileForm.Show
    
    ' �I�����ꂽ�t�@�C�������擾
    selectedFile = fileForm.GetSelectedFileName
    
    ' �t���p�X�𐶐�
    If selectedFile <> "" Then
        fullPath = folderPath & selectedFile
        
        ' Word���\���ŊJ��
        Set wordApp = CreateObject("Word.Application")
        wordApp.Visible = False ' Word���\���ŊJ��
        Set wordDoc = wordApp.Documents.Open(fullPath)
        
        ' Word��������{{}}�ň͂܂ꂽ�e�L�X�g���������Ēu��
        Do
            ' {{}}�ň͂܂ꂽ�e�L�X�g������
            Set rng = wordDoc.Content
            rng.Find.ClearFormatting
            rng.Find.text = "\{\{*\}\}"
            rng.Find.MatchWildcards = True ' ���C���h�J�[�h������L���ɂ���
            
            If rng.Find.Execute Then
                ' ���������e�L�X�g���擾
                findText = rng.text
                ' ���[�U�[�ɒu����̃e�L�X�g��q�˂�
                replacementText = InputBox("�u������e�L�X�g����͂��Ă�������:" & vbCrLf & "���������e�L�X�g: " & findText, "�e�L�X�g�̒u��")
                
                If replacementText <> "" Then
                    ' �e�L�X�g��u��
                    rng.text = replacementText
                Else
                    MsgBox "�u����̃e�L�X�g�����͂���܂���ł����B"
                End If
            Else
                Exit Do ' ����{{}}�ň͂܂ꂽ�e�L�X�g��������Ȃ������ꍇ�͏I��
            End If
        Loop
        
        ' �ύX���e���R�s�[���ăN���b�v�{�[�h�ɕۑ�
        wordDoc.Content.Copy ' �R�s�[���邱�ƂŃN���b�v�{�[�h�ɓ��e��ۑ�
        
        ' Outlook�𑀍삷�邽�߂̃I�u�W�F�N�g���쐬
        Set outlookApp = CreateObject("Outlook.Application")
        
        ' ���݃A�N�e�B�u�ȃ��[���A�C�e���܂��̓A�|�C���g�����g�A�C�e�����擾
        On Error Resume Next
        Set mailItem = outlookApp.ActiveInspector.CurrentItem
        Set appointmentItem = outlookApp.ActiveInspector.CurrentItem
        On Error GoTo 0
        
        ' ���[���A�C�e�����A�N�e�B�u�ȏꍇ
        If Not mailItem Is Nothing Then
            ' �N���b�v�{�[�h�̓��e��\��t���i������ێ��j
            mailItem.HTMLBody = mailItem.HTMLBody & "<br>" ' HTML�{���ɉ��s��ǉ�
            mailItem.GetInspector.CommandBars.ExecuteMso "Paste" ' �\��t��
            MsgBox "���[���Ƀe�L�X�g��\��t���܂����B"
        ' �A�|�C���g�����g�A�C�e�����A�N�e�B�u�ȏꍇ
        ElseIf Not appointmentItem Is Nothing Then
            ' �N���b�v�{�[�h�̓��e��\��t���i������ێ��j
            appointmentItem.Body = appointmentItem.Body & vbCrLf ' �{���ɉ��s��ǉ�
            appointmentItem.GetInspector.CommandBars.ExecuteMso "Paste" ' �\��t��
            MsgBox "�A�|�C���g�����g�Ƀe�L�X�g��\��t���܂����B"
        Else
            MsgBox "�A�N�e�B�u�ȃ��[���܂��̓A�|�C���g�����g��������܂���B"
        End If
        
        ' Word��ۑ������ɏI��
        wordDoc.Close False
        wordApp.Quit
    Else
        MsgBox "�t�@�C�����I������܂���ł����B"
    End If
End Sub

