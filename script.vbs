Option Explicit

Dim objFSO, WshShell, args, mode, inputFile, objFile, jsonText, json
Dim mainUrl
Dim i, paragraph

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
Set args = WScript.Arguments

If args.Count <> 1 Then
    WScript.Echo "Usage: cscript script.vbs input.json"
    WScript.Quit 1
End If

inputFile = args(0)

If Not objFSO.FileExists(inputFile) Then
    WScript.Echo "Input file not found: " & inputFile
    WScript.Quit 1
End If

' JSONパース関数（JScriptエンジンを利用）
Function ParseJSON(jsonString)
    Dim scriptEngine, parsed
    Set scriptEngine = CreateObject("MSScriptControl.ScriptControl")
    scriptEngine.Language = "JScript"
    scriptEngine.AddCode "function parseJSON(str){return eval('(' + str + ')');}"
    Set parsed = scriptEngine.Run("parseJSON", jsonString)
    Set ParseJSON = parsed
End Function

' input.jsonを読み込む
Set objFile = objFSO.OpenTextFile(inputFile, 1)
jsonText = objFile.ReadAll
objFile.Close

Set json = ParseJSON(jsonText)

mode = LCase(CStr(json("mode")))

If mode = "record" Then
    ' 記録モード
    If Not json.Exists("url") Then
        WScript.Echo "URL is not defined in input JSON."
        WScript.Quit 1
    End If

    mainUrl = CStr(json("url"))

    Dim outFile
    Set outFile = objFSO.CreateTextFile("RecordedActions.txt", True)

    outFile.WriteLine "操作開始: " & Now
    outFile.WriteLine "URL: " & mainUrl

    ' Chrome起動
    WshShell.Run "chrome.exe """ & mainUrl & """"
    WScript.Sleep 3000 ' ページの読み込み待ち

    Dim paragraphs
    Set paragraphs = json("paragraphs")

    For i = 0 To paragraphs.Count - 1
        paragraph = paragraphs.Item(i)
        Dim pText, imagePath, captionText
        pText = paragraph("text")
        imagePath = paragraph("imagePath")
        captionText = paragraph("caption")

        If pText <> "" Then
            outFile.WriteLine "段落: " & pText
            WshShell.SendKeys pText
            WshShell.SendKeys "{ENTER}"
            WScript.Sleep 500
        End If

        If imagePath <> "" Then
            If objFSO.FileExists(imagePath) Then
                outFile.WriteLine "画像: " & imagePath
                WshShell.SendKeys "%U" ' Alt+Uを想定
                WScript.Sleep 2000
                WshShell.SendKeys imagePath
                WshShell.SendKeys "{ENTER}"
                WScript.Sleep 2000

                outFile.WriteLine "キャプション: " & captionText
                WshShell.SendKeys captionText
                WshShell.SendKeys "{TAB}"
                WshShell.SendKeys captionText
                WshShell.SendKeys "{ENTER}"
                WScript.Sleep 1000
            Else
                WScript.Echo "指定された画像が見つかりません: " & imagePath
                outFile.WriteLine "画像: アップロード失敗"
            End If
        Else
            outFile.WriteLine "画像: アップロードなし"
        End If
    Next

    outFile.WriteLine "操作終了: " & Now
    outFile.Close

    WScript.Echo "操作をRecordedActions.txtに記録しました。"

ElseIf mode = "play" Then
    ' 再生モード
    If Not objFSO.FileExists("RecordedActions.txt") Then
        WScript.Echo "RecordedActions.txtが見つかりません。"
        WScript.Quit 1
    End If

    Set objFile = objFSO.OpenTextFile("RecordedActions.txt", 1)

    Dim line, paragraphText, imagePath, captionText
    Do Until objFile.AtEndOfStream
        line = objFile.ReadLine

        If InStr(line, "URL: ") > 0 Then
            mainUrl = Replace(line, "URL: ", "")
            WshShell.Run "chrome.exe """ & mainUrl & """"
            WScript.Sleep 3000
        ElseIf InStr(line, "段落: ") > 0 Then
            paragraphText = Replace(line, "段落: ", "")
            WshShell.SendKeys paragraphText
            WshShell.SendKeys "{ENTER}"
            WScript.Sleep 500
        ElseIf InStr(line, "画像: ") > 0 Then
            imagePath = Replace(line, "画像: ", "")
            If imagePath <> "アップロードなし" And imagePath <> "アップロード失敗" Then
                If objFSO.FileExists(imagePath) Then
                    WshShell.SendKeys "%U"
                    WScript.Sleep 2000
                    WshShell.SendKeys imagePath
                    WshShell.SendKeys "{ENTER}"
                    WScript.Sleep 2000

                    If Not objFile.AtEndOfStream Then
                        line = objFile.ReadLine
                        If InStr(line, "キャプション: ") > 0 Then
                            captionText = Replace(line, "キャプション: ", "")
                            WshShell.SendKeys captionText
                            WshShell.SendKeys "{TAB}"
                            WshShell.SendKeys captionText
                            WshShell.SendKeys "{ENTER}"
                            WScript.Sleep 1000
                        End If
                    End If
                Else
                    WScript.Echo "画像ファイルが見つかりません: " & imagePath
                End If
            End If
        End If
    Loop

    objFile.Close
    WScript.Echo "再生が完了しました。"

Else
    WScript.Echo "modeは'record'または'play'を指定してください。"
    WScript.Quit 1
End If
