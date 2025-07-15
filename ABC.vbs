Dim objExcel, objWorkbook, objSheet
Dim objWord, objDoc
Dim strExcelPath, strWordTemplatePath, strWordOutputPath
Dim i
Dim contentArrayB(49) ' 用于存储 B 列数据
Dim contentArrayC(49) ' 用于存储 C 列数据
' 定义文件路径
strExcelPath = "D:\workspace\VB\N2.xlsm"
strWordTemplatePath = "D:\workspace\VB\word.docx"
strWordOutputPath = "D:\workspace\VB\word0716.docx"
' 创建 Excel 应用程序对象
Set objExcel = CreateObject("Excel.Application")
' 打开 Excel 文件和工作表
Set objWorkbook = objExcel.Workbooks.Open(strExcelPath)
Set objSheet = objWorkbook.Sheets("word")
' 从列 B、C 中读取数据，共50项
For i = 0 To 49
    contentArrayB(i) = objSheet.Cells(i + 2, 2).Value ' 从 B 列的第 2 行开始
    contentArrayC(i) = objSheet.Cells(i + 2, 3).Value ' 从 C 列的第 2 行开始
Next
' 关闭 Excel 工作簿
objWorkbook.Close False
objExcel.Quit
' 创建 Word 应用程序对象并打开模板
Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Open(strWordTemplatePath)
objWord.Visible = False ' 如果不需要显示 Word 窗口，可以设置为 False
' 替换 Word 文档中的占位符
For i = 0 To 49
    ' 替换 B 列数据对应的占位符
    With objDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "<<AA" & (i + 1) & ">>"
        .Replacement.Text = contentArrayB(i)
        .Forward = True
        .Wrap = 1 ' wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute , , , , , , , , , , 2 ' 位置参数方式执行替换
    End With
    
    ' 替换 C 列数据对应的占位符
    With objDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "<<BB" & (i + 1) & ">>"
        .Replacement.Text = contentArrayC(i)
        .Forward = True
        .Wrap = 1 ' wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute , , , , , , , , , , 2 ' 位置参数方式执行替换
    End With
Next
' 保存修改后的 Word 文档
objDoc.SaveAs2 strWordOutputPath
objDoc.Close False
MsgBox "okkkkkk"

' 清理对象
objWord.Quit
Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
Set objDoc = Nothing
Set objWord = Nothing
