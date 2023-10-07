Option Explicit

    '定义全局变量，用于存储目标样式及对应RGB字体颜色
    Public stylesRgb As Variant
    Public g_bFinish As Boolean
    Public g_b黄云Flag As Boolean


Sub InitStylesRgb()
    '定义二维数组，用于存储目标样式及对应RGB字体颜色，并将其赋给全局变量stylesRgb
    Dim styles(6, 2) As Variant
    
    If g_b黄云Flag = True Then
        'styles(0, 0) = "正文":: styles(0, 1) = RGB(0, 0, 0)
        'styles(1, 0) = "标题 3":: styles(1, 1) = vbNull
        'styles(2, 0) = "Command":: styles(2, 1) = RGB(128, 0, 0)
        'styles(3, 0) = "标题 2":: styles(3, 1) = vbNull
    Else
        styles(0, 0) = "IPMI_正文":: styles(0, 1) = RGB(0, 0, 0)
        styles(1, 0) = "IPMI_二级标题":: styles(1, 1) = vbNull
        styles(2, 0) = "IPMI_正文":: styles(2, 1) = RGB(128, 0, 0)
        styles(3, 0) = "IPMI_【标题】":: styles(3, 1) = vbNull
    End If
    stylesRgb = styles
End Sub

'获取KeyWd在Sel中的位置
Function API_GetStartTruePos(Sel As Selection, KeyWd As String)
    Debug.Print "Func:API_GetStartTruePos:"; "Sel:"; Sel; ",KeyWd:"; KeyWd
    API_GetStartTruePos = Sel.Start + InStr(Sel.Text, KeyWd) - 1
End Function

'获取KeyWd在Sel中的位置
Function API_GetEndTruePos(Sel As Selection, KeyWd As String)
    Debug.Print "Func:  API_GetEndTruePos:"; "Sel:"; Sel; ",KeyWd:"; KeyWd
    API_GetEndTruePos = Sel.Start + InStr(Sel.Text, KeyWd)
End Function

'获取KeyWd在Sel中最后一个位置
Function API_GetEndRevTruePos(Sel As Selection, KeyWd As String)
    Debug.Print "Func:API_GetEndRevTruePos:"; "Sel:"; Sel; ",KeyWd:"; KeyWd
    API_GetEndRevTruePos = Sel.Start + InStrRev(Sel.Text, KeyWd)
End Function


Function API_GetSecondRangePos(Sel As Selection, KeyWd As String)
    Debug.Print "Func:  API_GetSecondRangePos:"; "Sel:"; Sel; ",KeyWd:"; KeyWd
    API_GetSecondRangePos = Sel.Start + InStr(Sel.Text, KeyWd)
End Function


'选中含【】整个内容
Sub API_GetNewPos_【】()
    Selection.SetRange Start:=API_GetStartTruePos(Selection, "【"), End:=API_GetEndTruePos(Selection, "】")
End Sub

'选中【】中的内容，不含括号
Sub API_GetNewPos_【】中的内容()
    Selection.SetRange Start:=API_GetStartTruePos(Selection, "【") + 1, End:=API_GetEndTruePos(Selection, "】") - 1
End Sub

'选中第一个】和最后一个【之间的内容
Sub API_GetNewPos_【】后面内容()
    Selection.SetRange Start:=API_GetStartTruePos(Selection, "】") + 2, End:=API_GetEndRevTruePos(Selection, "【") - 2
    Debug.Print "Range:"; Selection.Range
    
End Sub

Sub API_IPMI_捕获标题Selection()
    Debug.Print "Sub Test:"; "Selection:"; Selection.Text
    With Selection
     .StartIsActive = False
     .Extend Character:="】"
    End With
    Selection.SetRange Start:=Selection.Start, End:=Selection.End
    'Debug.Print "Sub API_IPMI_捕获标题Selection:"; "Selection3:"; Selection.Text
    Call_SS_标题【xxx】
End Sub

Sub API_IPMI_捕获内容Selection()
    'Debug.Print "Sub Test:"; "Selection:"; Selection.Text
    With Selection
     .StartIsActive = False
     .Extend Character:="【"
    End With
    Selection.SetRange Start:=Selection.Start + 1, End:=Selection.End - 1
    'Debug.Print "Sub API_IPMI_捕获内容Selection:"; "Selection3:"; Selection.Text
End Sub

Sub Test2()
    Debug.Print "Sub Test2"
    With Selection
        ' Collapse current selection to insertion point.
        .Collapse
        Debug.Print "1:"; Selection.Text
        ' Turn extend mode on.
        .Extend
        Debug.Print "2:"; Selection.Text
        ' Extend selection to word.
        .Extend
        Debug.Print "3:"; Selection.Text
        ' Extend selection to sentence.
        .Extend
        Debug.Print "4:"; Selection.Text
        '.ExtendMode = off
    End With
End Sub

'未完成
Sub API_IPMI_CmpKeyWd()
    Debug.Print "Sub API_IPMI_CmpKeyWd"
    Dim MyKeyWdArr
    MyKeyWdArr = Array("【命令】", "【参数说明】", "【所属权限模块】", "【支持产品】", "【返回值】", "【举例】")
    Dim KeyWd As Variant
    For Each KeyWd In MyKeyWdArr
        If InStr(1, Selection, KeyWd) Then
            'MsgBox Selection
            Debug.Print "Cmp find:"; Selection
            Exit For
        End If
    Next KeyWd
End Sub





Sub SS_FMT(style As String)
    Dim colorRGB As Variant
    Dim targetStyle As String
    Dim col As Integer
    Dim row As Integer
    Dim rng As Range
    Dim bIsFound As Integer
    
    bIsFound = 0

    '在全局变量中查找目标样式，并将对应的RGB字体颜色存储到colorRGB变量中
    For col = 0 To UBound(stylesRgb, 1)
        targetStyle = stylesRgb(col, 0)
        'MsgBox "loop: " & targetStyle & "and " & style
        If targetStyle = style Then
            '找到目标样式，标志位设置为1
            bIsFound = 1
        
            '找到的目标样式，无颜色设置要求
            If stylesRgb(col, 1) = vbNull Then
                'MsgBox "loop: " & "found no, array:" & colorRGB
            Else
                colorRGB = stylesRgb(col, 1)
                'MsgBox "loop: " & "found yes, array:" & colorRGB
                bIsFound = 2
            End If
            Exit For
        End If
    Next col
    
    '如果找到目标样式，则对选中区域应用该样式和RGB字体颜色
    
    Selection.ClearFormatting
    Selection.Font.Italic = False
    Selection.Font.Bold = False
    Selection.Font.Color = RGB(255, 255, 255)
    
    If bIsFound = 2 Then '设置样式和文字颜色
        Set rng = Selection.Range
        'MsgBox "Set style and color:[" & style & "][" & colorRGB & "]"
        With rng
            Debug.Print "style:"; style
            .style = style
            .Font.Color = colorRGB
            .Font.Italic = False
            .Font.Bold = False
        End With
    ElseIf bIsFound = 1 Then '设置样式
        Set rng = Selection.Range
        'MsgBox "Set style only:[" & style & "][" & colorRGB & "]"
        With rng
            .style = style
            .Font.Italic = False
            .Font.Bold = False
        End With
    Else
        '否则提示未找到目标样式
        MsgBox "无法找到指定的样式：" & style
    End If
    
End Sub


'调用函数，将选中区域样式设定为"正文"
Sub Call_SS_正文()
    '调用初始化函数，初始化全局变量stylesRgb
    InitStylesRgb
    
    If g_b黄云Flag = True Then
        SS_FMT "正文"
    Else
        SS_FMT "IPMI_正文"
    End If
End Sub


Sub Call_SS_中标题()
    '调用初始化函数，初始化全局变量stylesRgb
    InitStylesRgb
    
    If g_b黄云Flag = True Then
        SS_FMT "标题 3，heading 3，标题 3 Char1 Char，标题 3 Char1，标题 3 Char"
    Else
        SS_FMT "IPMI_二级标题"
    End If
End Sub

Sub Call_SS_大标题()
    InitStylesRgb
    
    If g_b黄云Flag = True Then
        SS_FMT "Command"
    Else
        SS_FMT "IPMI_一级标题"
    End If
End Sub

Sub Call_SS_标题【xxx】()
    InitStylesRgb
    If g_b黄云Flag = True Then
        SS_FMT "Command"
    Else
        SS_FMT "IPMI_【标题】"
    End If
End Sub

Sub Call_SS_表格()
    Dim tbl As Table
    Dim firstRow As row
    Dim nextRow As row

    '检查是否选中了一个表格
    If Selection.Information(wdWithInTable) = False Then
        MsgBox "请选中一个表格"
        Exit Sub
    End If

    'MsgBox "tbl---:" + vbLf + Selection.Text
    Set tbl = Selection.Tables(1)
    Set firstRow = tbl.Rows(1)
    
    '设置第一行的样式为"Table Heading"
    
    If g_b黄云Flag = True Then
        firstRow.Range.style = "正文"
    Else
        firstRow.Range.style = "IPMI_正文"
    End If

    firstRow.Range.Font.Color = RGB(0, 0, 0)
    firstRow.Range.Font.Italic = False
    
    firstRow.Range.Font.Name = "宋体"
    firstRow.Range.Font.Size = 11
    If g_b黄云Flag = True Then
        firstRow.Range.style = "Table Heading"
    Else
        firstRow.Range.style = "IPMI_表格头"
    End If

    '循环设置表格行的样式为"Table Text"
    For Each nextRow In tbl.Rows
        If nextRow.Index > 1 Then
            If g_b黄云Flag = True Then
                nextRow.Range.style = "正文"
            Else
                nextRow.Range.style = "IPMI_正文"
            End If
            nextRow.Range.Font.Color = RGB(0, 0, 0)
            nextRow.Range.Font.Italic = False
            nextRow.Range.Font.Bold = False
            nextRow.Range.Font.Name = "宋体"
            nextRow.Range.Font.Size = 11
            
            If g_b黄云Flag = True Then
                nextRow.Range.style = "Table Text"
            Else
                nextRow.Range.style = "IPMI_表格体"
            End If
        End If
    Next nextRow

    'MsgBox"表格样式设置完成"
    
    '表格不能使用通用的EndKey和MoveRight，无换行时，需要先MoveDown，再HomeKey到行首，最后MoveLeft，若有换行只需要MoveDown，这里手动处理比较方便，代码不知道怎么写
    Selection.MoveDown
	Selection.HomeKey
    Selection.MoveLeft
End Sub

Sub Call_SS_命令内容()
    Dim keywords() As Variant
    Dim selectedText As Range
    Dim keyword As Variant
    
    '定义关键词数组
    keywords = Array("connect_type", "hostname", "username", "password")
    
    Call_SS_正文
    
    '获取选中的文本范围
    Set selectedText = Selection.Range
    selectedText.Font.Name = "宋体"
    selectedText.Font.Size = 11

    '循环遍历关键词数组
    For Each keyword In keywords
        '重新选中文本范围，避免受之前的格式设置影响
        'MsgBox "keyword:" & keyword
        Set selectedText = Selection.Range
        'MsgBox "selectedText:" & selectedText
        selectedText.Select

        '使用高效的匹配方法
        With selectedText.Find
            .Text = keyword
            .Execute
        End With

        '如果匹配成功，则将关键词的字体设置为斜体
        If selectedText.Find.Found Then
            selectedText.Font.Italic = True
            selectedText.Find.Execute
        End If
    Next keyword
End Sub

Sub Call_SS_举例内容()
    Dim selText As Range
    Dim commandPos As Long
    Dim selectLen As Long
    
    
    '检查是否有选中文本
    If Selection.Type = wdSelectionIP Then
        MsgBox "请先选择需要格式化的文本。"
        Exit Sub
    End If
    
    Call_SS_正文

    '将选中文本赋值给Range对象
    Set selText = Selection.Range '每次使用前需要重新获取
    
    '检查文本中是否包含"Command"关键词
    commandPos = InStr(1, selText.Text, "COMMAND>")
    If commandPos = 0 Then
        MsgBox "所选文本中未找到关键词COMMAND>。"
        Exit Sub
    End If

    '获取选中区域总长度
    selectLen = selText.Characters.Count


    '设置关键词前面的文本格式
    selText.MoveStart wdCharacter, -1 '偏移到选中区域头部
    selText.MoveEnd wdCharacter, -(selectLen - commandPos + 1) '根据选中文字总长度，关键词位置，即获得去掉关键词及后面的
    selText.Font.Bold = False
    selText.Font.Italic = False

    'MsgBox "select1:" & selText
    Set selText = Selection.Range '每次使用前需要重新获取
    'MsgBox "select2:" & selText & "post:" & commandPos & " selectLen:" & selectLen
    '设置关键词自身的文本格式
    selText.MoveStart wdCharacter, commandPos - 1 '光标移动到首次出现关键词的地方
    selText.MoveEnd wdCharacter, -(selectLen - commandPos - 9) '根据当前选中内容的总长度，减去需加粗的区域内容，并进行选中，
    'MsgBox "select3:" & selText
    selText.Font.Name = "Courier New"
    selText.Font.Bold = True
    selText.Font.Italic = False
    selText.Font.Size = 8.5
    
    Set selText = Selection.Range   '每次使用前需要重新获取
    '设置关键词后面的文本格式
    selText.MoveEnd wdCharacter, 0  '将Range对象的结束位置设置为选中范围的末尾
    selText.MoveStart wdCharacter, commandPos + 9 '将Range对象的起始位置设置为关键词处
    selText.Font.Name = "Courier New"
    selText.Font.Bold = False
    selText.Font.Size = 8.5
End Sub

Sub API_IPMI_TableTrimAndSelect()
    Dim TblLoop As Long
    TblLoop = 0
    
    If Selection.Information(wdWithInTable) Then
        Selection.Tables(1).Select
        'MsgBox "after3" & vbLf & Selection
        Debug.Print "Table set wdWithInTable？ ->Find table"
    Else
        Debug.Print "Table set wdWithInTable？ ->No find, try trim"
        Do
            '从末尾去掉内容，直到选中了表格
            Selection.SetRange Start:=Selection.Start, End:=Selection.End - 1
             TblLoop = TblLoop + 1
             Debug.Print "Do reset range, to fit table, TblLoop:"; TblLoop
        Loop Until Selection.Information(wdWithInTable) = True Or TblLoop > 10
    End If
End Sub


Sub API_IPMI_CmpKeyWdEach()
    Debug.Print "Sub API_IPMI_CmpKeyWdEach"

    API_IPMI_捕获标题Selection
    'MsgBox Selection
    If InStr(1, Selection, "【命令】") Then
        'MsgBox Selection
        Selection.MoveEndUntil Cset:="【", Count:=wdForward
        Selection.SetRange Start:=(Selection.Start + Len("【命令】") + 1), End:=Selection.End
        Call_SS_命令内容
    ElseIf InStr(1, Selection, "【参数说明】") Then
        Selection.MoveEndUntil Cset:="字", Count:=wdForward
        Selection.SetRange Start:=(Selection.Start + Len("【参数说明】") + 1), End:=Selection.End
        Selection.MoveEndUntil Cset:="【", Count:=wdForward
        Dim TblLoop As Long
        TblLoop = 0
        
        If Selection.Information(wdWithInTable) Then
            Selection.Tables(1).Select
            'MsgBox "after3" & vbLf & Selection
            Debug.Print "【参数说明】wdWithInTable？ ->Find table"
        Else
            Debug.Print "【参数说明】wdWithInTable？ ->No find, try trim"
            Do
                '从末尾去掉内容，直到选中了表格
                Selection.SetRange Start:=Selection.Start, End:=Selection.End - 1
                 TblLoop = TblLoop + 1
                 Debug.Print "Do reset range, to fit table, TblLoop:"; TblLoop
            Loop Until Selection.Information(wdWithInTable) = True Or TblLoop > 10
        End If
        Call_SS_表格
    ElseIf InStr(1, Selection, "【所属权限模块】") Then
        'MsgBox Selection
        Selection.MoveEndUntil Cset:="【", Count:=wdForward
        Selection.SetRange Start:=(Selection.Start + Len("【所属权限模块】") + 1), End:=Selection.End
        Call_SS_正文
    ElseIf InStr(1, Selection, "【支持产品】") Then
        'MsgBox Selection
        Selection.MoveEndUntil Cset:="【", Count:=wdForward
        Selection.SetRange Start:=(Selection.Start + Len("【支持产品】") + 1), End:=Selection.End
        Call_SS_正文
    ElseIf InStr(1, Selection, "【返回值】") Then
        'MsgBox Selection
        Selection.MoveEndUntil Cset:="字", Count:=wdForward
        Selection.SetRange Start:=(Selection.Start + Len("【返回值】") + 1), End:=Selection.End
        API_IPMI_TableTrimAndSelect
        Call_SS_表格
    ElseIf InStr(1, Selection, "【举例】") Then
        'MsgBox Selection
        Selection.MoveEndUntil Cset:="【", Count:=wdForward
        Selection.SetRange Start:=(Selection.Start + Len("【举例】") + 1), End:=Selection.End
        Call_SS_举例内容
    ElseIf InStr(1, Selection, "【修改记录】") Then
        'MsgBox Selection
        Selection.MoveEndUntil Cset:="【", Count:=wdForward
        Selection.SetRange Start:=(Selection.Start + Len("【修改记录】") + 1), End:=Selection.End
        Call_SS_正文
	ElseIf InStr(1, Selection, "【修改历史】") Then
        'MsgBox Selection
        Selection.MoveEndUntil Cset:="【", Count:=wdForward
        Selection.SetRange Start:=(Selection.Start + Len("【修改历史】") + 1), End:=Selection.End
        Call_SS_正文
	ElseIf InStr(1, Selection, "【其他说明】") Then
        'MsgBox Selection
        Selection.MoveEndUntil Cset:="【", Count:=wdForward
        Selection.SetRange Start:=(Selection.Start + Len("【其他说明】") + 1), End:=Selection.End
        Call_SS_正文
    ElseIf InStr(1, Selection, "【End】") Then
        'Selection.MoveEndUntil Cset:="【", Count:=wdForward
        'Selection.SetRange Start:=(Selection.Start + Len("【End】") + 1), End:=Selection.End
        'MsgBox Selection
        Selection.Delete
        'MsgBox "Find 【End】"
        g_bFinish = True
    End If
End Sub

Sub API_IPMI_RunStyleFormat()
    g_bFinish = False '初始化中止变量，检测到完成条件时，会退出死循环
    g_b黄云Flag = False
    
    Dim Counter As Long
    Counter = 0 '避免死循环
    
    Dim TotalLen As Long
    TotalLen = Selection.StoryLength
    Debug.Print "TotalLen:"; TotalLen
    
    Dim StartStoryPos As Long
    StartStoryPos = Selection.Start
    Debug.Print "StartStoryPos:"; StartStoryPos
        
    '选中的字段末尾添加新一行的结束符【End】，用于判定
    Selection.InsertAfter vbLf & "【End】"

    '调整目标起始偏移位置
    Selection.HomeKey
    'MsgBox "Into do"
    Do
        
        '匹配条件检测，并进行格式化调整
        API_IPMI_CmpKeyWdEach
        
        '执行复位动作
        Selection.EndKey
        Selection.MoveRight
        
        Counter = Counter + 1
        If Counter > 10 Then
            g_bFinish = True
            MsgBox "Error Counter=" & Counter & ", please check each box"
        End If
    Loop Until g_bFinish = True
    
    '完成删除结束符【End】
End Sub

