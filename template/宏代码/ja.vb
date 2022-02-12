'試験用紙を生成する
Sub generate()
    ActiveSheet.Unprotect
    Application.ScreenUpdating = False 'スクリーンリフレッシュを停止する
    Dim curRow As Integer '該当行番
    curRow = 4 '開始行番
    'データ列の定義
    Const descriptionPos As String = "G1"
    Dim description As String
    Const authorPos As String = "V1"
    Dim author As String
    Const scoreCol As String = "B"
    Dim score As String
    Const typeCol As String = "C"
    Dim sType As String
    Const questionCol As String = "D"
    Dim question As String
    Const opt1Col As String = "AB"
    Dim opt1 As String
    Const opt2Col As String = "AE"
    Dim opt2 As String
    Const opt3Col As String = "AH"
    Dim opt3 As String
    Const opt4Col As String = "AK"
    Dim opt4 As String
    Const commentCol As String = "AN"
    Dim comment As String
    Const imageCol As String = "BB"
    Dim image As String
    Const audioCol As String = "BP"
    Dim audio As String
    Const videoCol As String = "CD"
    Dim video As String
    'エラーメッセージをクリアする
    setWhiteBg descriptionPos
    setWhiteBg authorPos
    setGrayBg "B4:C104"
    setWhiteBg "D4:D103"
    setGrayBg "AB4:CD103"
    'JSONを初期化する
    Dim data As String
    data = "{"
    'ワークシートを選択する
    With Worksheets("Sheet1")
        '試験説明
        description = .Range(descriptionPos).Value
        If description = "" Then
            setRedBg descriptionPos
            MsgBox "【試験説明】に内容を入力してください。"
            reset
            Exit Sub
        End If
        data = data & qt("description") & ":" & qt(description) & ","
        '出題者
        author = .Range(authorPos).Value
        If author = "" Then
            setRedBg authorPos
            MsgBox "【出題者】に名前を入力してください。"
            reset
            Exit Sub
        End If
        data = data & qt("author") & ":" & qt(author) & ","
        '今の時間
        data = data & qt("datetime") & ":" & qt(datetimeNow) & ","
        '問題内容
        data = data & qt("content") & ":["
        While .Range(questionCol & curRow).Value <> ""
            data = data & "{"
            '問題
            question = .Range(questionCol & curRow).Value
            data = data & qt("question") & ":" & qt(question) & ","
            'タイプ
            sType = .Range(typeCol & curRow).Value
            If sType = "" Then
                setRedBg typeCol & curRow
                MsgBox "タイプ情報を選択してください。"
                reset
                Exit Sub
            End If
            If sType = "空欄" Then
                data = data & qt("type") & ":" & qt("text") & ","
            ElseIf sType = "選択" Then
                data = data & qt("type") & ":" & qt("radio") & ","
            ElseIf sType = "主観" Then
                data = data & qt("type") & ":" & qt("textarea") & ","
            End If
            '点数
            score = .Range(scoreCol & curRow).Value
            If score = "" Then
                setRedBg scoreCol & curRow
                MsgBox "点数情報を入力してください。"
                reset
                Exit Sub
            End If
            data = data & qt("score") & ":" & qt(score) & ","
            '画像
            image = .Range(imageCol & curRow).Value
            data = data & qt("image") & ":" & qt(image) & ","
            '音声
            audio = .Range(audioCol & curRow).Value
            data = data & qt("audio") & ":" & qt(audio) & ","
            'ビデオ
            video = .Range(videoCol & curRow).Value
            data = data & qt("video") & ":" & qt(video) & ","
            '選択肢
            If sType = "選択" Then
                opt1 = .Range(opt1Col & curRow).Value
                If opt1 = "" Then
                    setRedBg opt1Col & curRow
                    MsgBox "選択肢1を入力してください。"
                    reset
                    Exit Sub
                End If
                opt2 = .Range(opt2Col & curRow).Value
                If opt2 = "" Then
                    setRedBg opt2Col & curRow
                    MsgBox "選択肢2を入力してください。"
                    reset
                    Exit Sub
                End If
                opt3 = .Range(opt3Col & curRow).Value
                If opt3 = "" Then
                    setRedBg opt3Col & curRow
                    MsgBox "選択肢2を入力してください。"
                    reset
                    Exit Sub
                End If
                opt4 = .Range(opt4Col & curRow).Value
                If opt4 = "" Then
                    setRedBg opt4Col & curRow
                    MsgBox "選択肢4を入力してください。"
                    reset
                    Exit Sub
                End If
                data = data & qt("options") & ":[" & qt(opt1) & "," & qt(opt2) & "," & qt(opt3) & "," & qt(opt4) & "],"
            Else
                '選択じゃない問題
                data = data & qt("options") & ":[],"
            End If
            '備考（書取・翻訳）
            comment = .Range(commentCol & curRow).Value
            data = data & qt("comment") & ":" & qt(comment) & ","
            '次の問題
            data = Left(data, Len(data) - 1) & "},"
            curRow = curRow + 1
        Wend
        '全部の問題内容
        data = Left(data, Len(data) - 1) & "]"
    End With
    data = data & "}"
    'JSONに書き込む
    save data
    '復旧
    reset
    Range("A1").Select
    MsgBox "エクスポート成功"
End Sub

'復旧
Sub reset()
    Application.ScreenUpdating = True 'スクリーンリフレッシュを再開する
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

'JSONに書き込む
Sub save(data)
    Dim filePath As String
    filePath = ThisWorkbook.Path & Application.PathSeparator & "data.json"
    
    'エクスポート
    ' Open filePath For Output As #1
    ' Print #1, data
    ' Close #1

    'エクスポート（UTF-8）
    Dim outStream As Object
    Set outStream = CreateObject("ADODB.Stream")
    With outStream
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText data
        .SaveToFile filePath, 2
    End With
End Sub

'改行
Function nl(text)
    nl = Chr(13) & Chr(10) & text
End Function

'コーテーションマーク
Function qt(text)
    qt = Chr(34) & text & Chr(34)
End Function

'白い背景を設定する
Sub setWhiteBg(rng)
    Range(rng).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
End Sub
    
'灰色の背景を設定する
Sub setGrayBg(rng)
    Range(rng).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
End Sub

'赤い背景を設定する
Sub setRedBg(rng)
    Range(rng).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
End Sub

'今の日付と時間
Function datetimeNow()
    Dim datetime As String
    '年
    datetime = Year(Now) & "/"
    '月
    If Month(Now) < 10 Then
        datetime = datetime & "0"
    End If
    datetime = datetime & Month(Now) & "/"
    '日
    If Day(Now) < 10 Then
        datetime = datetime & "0"
    End If
    datetime = datetime & Day(Now) & " "
    '時
    If Hour(Now) < 10 Then
        datetime = datetime & "0"
    End If
    datetime = datetime & Hour(Now) & ":"
    '分
    If Minute(Now) < 10 Then
        datetime = datetime & "0"
    End If
    datetime = datetime & Minute(Now) & ":"
    '秒
    If Second(Now) < 10 Then
        datetime = datetime & "0"
    End If
    datetime = datetime & Second(Now)
    datetimeNow = datetime
End Function
