'生成题库
Sub generate()
    ActiveSheet.Unprotect
    Application.ScreenUpdating = False '关闭屏幕刷新
    Dim curRow As Integer '当前行
    curRow = 4 '起始行
    '数据列定义
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
    '清空错误信息
    setWhiteBg descriptionPos
    setWhiteBg authorPos
    setGrayBg "B4:C104"
    setWhiteBg "D4:D103"
    setGrayBg "AB4:CD103"
    '初始化JSON
    Dim data As String
    data = "{"
    '选中数据表
    With Worksheets("Sheet1")
        '题库描述
        description = .Range(descriptionPos).Value
        If description = "" Then
            setRedBg descriptionPos
            MsgBox "请填写【题库描述】。"
            reset
            Exit Sub
        End If
        data = data & qt("description") & ":" & qt(description) & ","
        '出题人
        author = .Range(authorPos).Value
        If author = "" Then
            setRedBg authorPos
            MsgBox "请填写【出题人】。"
            reset
            Exit Sub
        End If
        data = data & qt("author") & ":" & qt(author) & ","
        '当前时间
        data = data & qt("datetime") & ":" & qt(datetimeNow) & ","
        '题目内容
        data = data & qt("content") & ":["
        While .Range(questionCol & curRow).Value <> ""
            data = data & "{"
            '题干
            question = .Range(questionCol & curRow).Value
            data = data & qt("question") & ":" & qt(question) & ","
            '题型
            sType = .Range(typeCol & curRow).Value
            If sType = "" Then
                setRedBg typeCol & curRow
                MsgBox "请填写题型信息。"
                reset
                Exit Sub
            End If
            If sType = "填空" Then
                data = data & qt("type") & ":" & qt("text") & ","
            ElseIf sType = "选择" Then
                data = data & qt("type") & ":" & qt("radio") & ","
            ElseIf sType = "主观" Then
                data = data & qt("type") & ":" & qt("textarea") & ","
            End If
            '分数
            score = .Range(scoreCol & curRow).Value
            If score = "" Then
                setRedBg scoreCol & curRow
                MsgBox "请填写分数信息。"
                reset
                Exit Sub
            End If
            data = data & qt("score") & ":" & qt(score) & ","
            '图片
            image = .Range(imageCol & curRow).Value
            data = data & qt("image") & ":" & qt(image) & ","
            '音频
            audio = .Range(audioCol & curRow).Value
            data = data & qt("audio") & ":" & qt(audio) & ","
            '视频
            video = .Range(videoCol & curRow).Value
            data = data & qt("video") & ":" & qt(video) & ","
            '选项
            If sType = "选择" Then
                opt1 = .Range(opt1Col & curRow).Value
                If opt1 = "" Then
                    setRedBg opt1Col & curRow
                    MsgBox "请填写选项1信息。"
                    reset
                    Exit Sub
                End If
                opt2 = .Range(opt2Col & curRow).Value
                If opt2 = "" Then
                    setRedBg opt2Col & curRow
                    MsgBox "请填写选项2信息。"
                    reset
                    Exit Sub
                End If
                opt3 = .Range(opt3Col & curRow).Value
                If opt3 = "" Then
                    setRedBg opt3Col & curRow
                    MsgBox "请填写选项3信息。"
                    reset
                    Exit Sub
                End If
                opt4 = .Range(opt4Col & curRow).Value
                If opt4 = "" Then
                    setRedBg opt4Col & curRow
                    MsgBox "请填写选项4信息。"
                    reset
                    Exit Sub
                End If
                data = data & qt("options") & ":[" & qt(opt1) & "," & qt(opt2) & "," & qt(opt3) & "," & qt(opt4) & "],"
            Else
                '非选择题
                data = data & qt("options") & ":[],"
            End If
            '备注（听写/翻译）
            comment = .Range(commentCol & curRow).Value
            data = data & qt("comment") & ":" & qt(comment) & ","
            '下一题
            data = Left(data, Len(data) - 1) & "},"
            curRow = curRow + 1
        Wend
        '全部题目内容
        data = Left(data, Len(data) - 1) & "]"
    End With
    data = data & "}"
    '写入JSON
    save data
    '复位
    reset
    Range("A1").Select
    MsgBox "导出成功。"
End Sub

'复位
Sub reset()
    Application.ScreenUpdating = True '恢复屏幕刷新
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

'写入JSON
Sub save(data)
    Dim filePath As String
    filePath = ThisWorkbook.Path & Application.PathSeparator & "data.json"
    
    '导出内容
    ' Open filePath For Output As #1
    ' Print #1, data
    ' Close #1

    '导出 UTF-8 编码
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

'另起新行
Function nl(text)
    nl = Chr(13) & Chr(10) & text
End Function

'外包双引号
Function qt(text)
    qt = Chr(34) & text & Chr(34)
End Function

'设置白色背景
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
    
'设置灰色背景
Sub setGrayBg(rng)
    Range(rng).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.0499893185216834
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
End Sub

'设置红色背景
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

'当前日期时间
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
    '时
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
