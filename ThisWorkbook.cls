VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub FindKeywordsInAllSheets()
    Dim wsTool As Worksheet
    Dim wsSearch As Worksheet
    Dim rngKeywords As Range
    Dim keyword As Range
    Dim result As String
    Dim count As Integer

    ' 获取 Tool 工作表
    Set wsTool = ThisWorkbook.Worksheets("Tool")

    ' 获取 A 列中的关键字范围
    Set rngKeywords = wsTool.Range("A:A").SpecialCells(xlCellTypeConstants)

    ' 循环遍历关键字
    For Each keyword In rngKeywords
        result = ""
        count = 0

        ' 循环遍历所有工作表
        For Each wsSearch In ThisWorkbook.Worksheets
            If wsSearch.Name <> wsTool.Name Then
                ' 在工作表中查找关键字
                ' 如果关键字存在，则将工作表名称添加到结果中
                If Not wsSearch.Range("A:A").Find(keyword.Value, LookIn:=xlValues, LookAt:=xlPart) Is Nothing Then
                    If count = 0 Then
                        result = wsSearch.Name
                    Else
                        result = result & ", " & wsSearch.Name
                    End If
                    count = count + 1
                End If
            End If
        Next wsSearch

        ' 将结果写入 Tool 工作表中
        keyword.Offset(0, 1).Value = result

    Next keyword

    ' 替换结果中的Sheet2为A，Sheet3为B
    For Each keyword In rngKeywords
        Dim replaceValue As String
        replaceValue = keyword.Offset(0, 1).Value
        replaceValue = Replace(replaceValue, "Sheet2", "2")
        replaceValue = Replace(replaceValue, "Sheet3", "3")
        replaceValue = Replace(replaceValue, "Sheet4", "4")
        keyword.Offset(0, 1).Value = replaceValue
    Next keyword
    
End Sub
