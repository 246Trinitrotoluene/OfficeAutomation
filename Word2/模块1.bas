Attribute VB_Name = "ģ��1"
Sub ���ΪTXT()
Dim DocxPath As String
Dim TextPath As String
Dim Desktop As String
Desktop = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"


' ���������Ӧ���
' GB 2312   936
' GB18030 54936
' BIG5      950
' UTF8    65001


    '������
    Selection.WholeStory
    Selection.Fields.Update
    
    
    '��ȡ�ļ�·���������������ΪTXT/UTF8
    DocxPath = ActiveDocument.FullName
    If InStr(1, DocxPath, ".txt", vbTextCompare) = 0 Then
        
        
        TextPath = ActiveDocument.Paragraphs(1).Range.Text
        TextPath = Left(TextPath, Len(TextPath) - 1) & ".txt"
        TextPath = Desktop & TextPath
        MsgBox (TextPath)
        
        
        ChangeFileOpenDirectory Desktop
        ActiveDocument.SaveAs2 filename:=TextPath, FileFormat:=wdFormatText, Encoding:=65001, _
            AddToRecentFiles:=False, AllowSubstitutions:=False, LineEnding:=wdCRLF
        
        
        '��Word�йر�TXT�ĵ�����DOCX�ĵ�
        ActiveDocument.Close 0
        Documents.Open filename:=DocxPath
        'ActiveDocument.Save
        
        
        '��TXT�ļ�
        Shell ("notepad " & TextPath)
        
        
    Else
        MsgBox ("����TXT�ļ�������ת��")
        
        
    End If
    
    
End Sub


