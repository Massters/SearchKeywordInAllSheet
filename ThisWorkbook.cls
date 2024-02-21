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

    ' ��ȡ Tool ������
    Set wsTool = ThisWorkbook.Worksheets("Tool")

    ' ��ȡ A ���еĹؼ��ַ�Χ
    Set rngKeywords = wsTool.Range("A:A").SpecialCells(xlCellTypeConstants)

    ' ѭ�������ؼ���
    For Each keyword In rngKeywords
        result = ""
        count = 0

        ' ѭ���������й�����
        For Each wsSearch In ThisWorkbook.Worksheets
            If wsSearch.Name <> wsTool.Name Then
                ' �ڹ������в��ҹؼ���
                ' ����ؼ��ִ��ڣ��򽫹������������ӵ������
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

        ' �����д�� Tool ��������
        keyword.Offset(0, 1).Value = result

    Next keyword
End Sub