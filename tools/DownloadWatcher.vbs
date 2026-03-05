Option Explicit

Dim sh, fso, target, lastCount
Set sh  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' ★ ユーザーごとの Downloads を自動解決（固定にしたい場合は直書きでOK）
target = sh.ExpandEnvironmentStrings("%USERPROFILE%\Downloads")

' ===== 開いているか判定する関数（重複オープン防止の要） =====
Function IsFolderOpen(folderPath)
    On Error Resume Next
    Dim shellApp, w, openPath, isOpen
    isOpen = False
    Set shellApp = CreateObject("Shell.Application")

    ' 参照用に正規化（末尾の \ を削除、大小無視で比較）
    folderPath = Trim(folderPath)
    If Right(folderPath, 1) = "\" Then folderPath = Left(folderPath, Len(folderPath)-1)

    For Each w In shellApp.Windows
        ' Explorer 以外（Edge/IE/Office 等）はスキップ
        If Not w Is Nothing Then
            ' フォルダビューのみ対象（Document が Folder を持つかで判断）
            If Not w.Document Is Nothing Then
                If Err.Number = 0 Then
                    ' 表示中フォルダの実パスを取得
                    openPath = ""
                    On Error Resume Next
                    openPath = w.Document.Folder.Self.Path
                    If Err.Number = 0 Then
                        ' 末尾 \ を削除してケース無視で比較
                        If Right(openPath, 1) = "\" Then openPath = Left(openPath, Len(openPath)-1)
                        If StrComp(openPath, folderPath, vbTextCompare) = 0 Then
                            isOpen = True
                            Exit For
                        End If
                    End If
                    On Error GoTo 0
                Else
                    Err.Clear
                End If
            End If
        End If
    Next

    IsFolderOpen = isOpen
    On Error GoTo 0
End Function
' ============================================================

' 初期ファイル数を取得（フォルダのみ除外）
Function CountFiles(path)
    Dim folder, file, c
    c = 0
    If fso.FolderExists(path) Then
        Set folder = fso.GetFolder(path)
        For Each file In folder.Files
            c = c + 1
        Next
    End If
    CountFiles = c
End Function

lastCount = CountFiles(target)

Do
    Dim current
    current = CountFiles(target)

    If current > lastCount Then
        lastCount = current
        ' ★ 既に開いていなければだけ開く
        If Not IsFolderOpen(target) Then
            sh.Run "explorer.exe """ & target & """", 1, False
        End If
    ElseIf current < lastCount Then
        lastCount = current
    End If

    WScript.Sleep 2000
Loop本スクリプトは、ファイルをダウンロードする(ダウンロードフォルダにファイルが入る)毎に
ダウンロードフォルダが開かれるものとなります

メリット：
/エクスプローラー→ダウンロードフォルダを開く、という手間が省かれる
/ダウンロード毎にフォルダが開くので、リマインドになる

使用方法：
watch_downloadフォルダ内の、create_shortcut.vbsをダブルクリックする

走査フロー：
PCにログインする毎にスタートアップフォルダ内のwatch.vbsショートカットを通じて、
watch.vbsが開き、以降ログオフまでバックグラウンドで動き続ける
そうしてダウンロードフォルダにファイルが増える毎にダウンロードフォルダを開く

仕様：
/既にダウンロードフォルダを開いている場合は開きません
/ダウンロードフォルダからファイルを削除した場合は開きません
/本スクリプトを停止したい場合はremove.vbsをダブルクリックしてください
 そうすると、次回ログイン以降本スクリプトは動きません