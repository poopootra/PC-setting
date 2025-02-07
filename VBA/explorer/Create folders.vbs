Set ws = CreateObject("WScript.Shell")
cd = ws.CurrentDirectory
rd = "..\..\"
md = rd & "\20 Meeting\"

DirArray = Array("\00 Proposal","\01 Project management","\02 Input","\03 Interview","\04 Analysis","\20 Meeting", "\21 Deliverable","\91 Email","\01 Project management\01 Project note","\01 Project management\02 Mail draft","\01 Project management\03 Deliverable template","\02 Input\01 External","\02 Input\02 Internal","\03 Interview\01 Discussion guide","\03 Interview\02 Interview note","\04 Analysis\01","\04 Analysis\02","\04 Analysis\03","\21 Deliverable\01","\21 Deliverable\02","\21 Deliverable\03","\91 Email\01 History","\91 Email\02 Template")

Set fso = CreateObject("Scripting.FileSystemObject")

For Each value In DirArray
    If Not (fso.FolderExists(rd & value)) Then
        ' フォルダの作成
        fso.CreateFolder (rd & value)
    End If
Next



' ループ処理
d_base = Date
IntDif = 2 - Weekday(d_base)
If IntDif <= 0 Then
    IntDif = IntDif + 7
End If
For i = 0 To 12
    d = DateAdd("d", d_base, IntDif + 7 * i)
    ' yyyy-mm-dd
    s1 = Year(d) & "-" & Right("00" & Month(d), 2) & "-" & Right("00" & Day(d), 2)
    ' 作成するフォルダのパスを生成
    mkdirPath = md & "\" & "Week of " & s1
    ' フォルダが存在しない場合のみフォルダを作成する
    If Not (fso.FolderExists(mkdirPath)) Then
        ' フォルダの作成
        fso.CreateFolder (mkdirPath)
    End If


Next


' オブジェクトの解放
Set fso = Nothing
Set ws = Nothing

MsgBox("Done")

WScript.Quit