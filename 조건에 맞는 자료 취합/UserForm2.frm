VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6456
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   6444
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If ComboBox1.ListIndex = -1 Then
        MsgBox ("ȣ�⸦ �����ϼ���")
        Exit Sub
    ElseIf ComboBox2.ListIndex = -1 Then
        MsgBox ("CS�з��� �����ϼ���")
        Exit Sub
    ElseIf ComboBox3.ListIndex = -1 Then
        MsgBox ("CDA�з��� �����ϼ���")
        Exit Sub
    ElseIf ComboBox4.ListIndex = -1 Then
        MsgBox ("NEI�з��� �����ϼ���")
        Exit Sub
    End If
    
    Dim SummarySheet As Worksheet
    Dim CSSheet As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim rngFound As Range
    Dim str1 As String
    Dim str2 As String
    Dim strFirst As String
    Dim CScolumn As Long
    Dim CDAcolumn As Long
    
    Set SummarySheet = Workbooks(ActiveWorkbook.Name).Worksheets(ComboBox1.Value)
    Set CSSheet = Workbooks(ActiveWorkbook.Name).Worksheets("CS�з�")
    
    LastRow = SummarySheet.Cells(Cells.Rows.Count, "b").End(xlUp).Row
    
    For i = 5 To LastRow
        'D�� E���� ���� ã�ƾ��� ��Ʈ������ ����
        str1 = SummarySheet.Cells(i, "D")
        str2 = SummarySheet.Cells(i, "E")
        
        
        
        'PBS�� �������а��� "CS�з�"��Ʈ�� ��Ī�ؼ� CS�� �˾Ƴ��� �� CS�� H,I,J,K 4�� ���� ����
        Set rngFound = CSSheet.Columns("F").Find(str1, Cells(2, "F"), xlValues, xlWhole)

        If Not rngFound Is Nothing Then
            strFirst = rngFound.Address
            Do
                If LCase(CSSheet.Cells(rngFound.Row, "D").Text) = LCase(str2) Then
                    'Found a match
                    'MsgBox "Found a match at: " & rngFound.Row & Chr(10) & "Value in column C: " & Cells(rngFound.Row, "C").Text & Chr(10) & "Value in column D: " & Cells(rngFound.Row, "D").Text
                    Exit Do
                End If
                Set rngFound = Columns("F").Find(str1, rngFound, xlValues, xlWhole)
            Loop While (Not rngFound Is Nothing)
        End If
        If rngFound Is Nothing Then GoTo continue
        
        Set rngFound = CSSheet.Rows(rngFound.Row).Find("O", , xlValues, xlWhole)
        If rngFound Is Nothing Then GoTo continue
        '���� �˾Ƴ� CS ���°� �޺��ڽ�2�� �з��� ��ġ�ϸ� �۾� ����
        Select Case ComboBox2.ListIndex
            Case 0
                CScolumn = 8
            Case 1
                CScolumn = 9
            Case 2
                CScolumn = 10
            Case 3
                CScolumn = 11
        End Select
        If CScolumn = rngFound.Column Then
            Set rngFound = SummarySheet.Rows(i).Find("O", Cells(i, 12), xlValues, xlWhole)
            If rngFound Is Nothing Then GoTo continue
            Select Case ComboBox3.ListIndex
                Case 0
                    CDAcolumn = 12
                Case 1
                    CDAcolumn = 13
                Case 2
                    CDAcolumn = 14
                Case 3
                    CDAcolumn = 15
                Case 4
                    CDAcolumn = 16
            End Select
            'SummarySheet�� CDA1-5 �� L,M,N,O,P 5�� ���� ����
            '���� �޺��ڽ�3�� �з��� ���� ���� üũ�� �Ǿ�������
            'AC���� �޺��ڽ�4�� ���õ� ���� �־��ش�
            If CDAcolumn = rngFound.Column Then
                SummarySheet.Range("AC" & i).Value = ComboBox4.Value
            End If
        End If
continue:
    Next
    
End Sub

Private Sub CommandButton2_Click()
    Dim SummarySheet As Worksheet
    Dim countCDASheet As Worksheet
    Dim FolderPath As String
    Dim SelectedFiles As Variant
    Dim NRow As Long
    Dim FileName As String
    Dim NFile As Long
    Dim WorkBk As Workbook
    Dim SourceRange As Range
    Dim DestRange As Range
    Dim LastRow As Long
    Dim answer As Integer
    Dim countCDA As Integer
    Dim countSecure As Integer
    Dim countNonSecure As Integer
    Dim countNon As Integer
    Dim i As Integer
    Dim Dest_LastRow
    Dim a As Integer
    Dim b As String
    
    If ComboBox1.ListIndex = -1 Then
        MsgBox ("ȣ�⸦ �����ϼ���")
        Exit Sub
    End If
    
    Set SummarySheet = Workbooks(ActiveWorkbook.Name).Worksheets(ComboBox1.Value)
    a = SummarySheet.Index
    b = SummarySheet.Name
    Set countCDASheet = Workbooks(ActiveWorkbook.Name).Worksheets("����")
    
    If b <> "����5ȣ��" And b <> "����6ȣ��" Then
        Exit Sub
    End If
    
    answer = MsgBox("���� �����->YesŬ�� / �̾� ���̱�->NoŬ��", vbYesNo + vbQuestion, "Sheet�ʱ�ȭ����")
    If answer = vbYes Then
        If b = "����5ȣ��" Then
            LastRow = countCDASheet.Cells(Cells.Rows.Count, "i").End(xlUp).Row
            countCDASheet.Range("i1", "m" & LastRow).ClearContents
        ElseIf b = "����6ȣ��" Then
            LastRow = countCDASheet.Cells(Cells.Rows.Count, "m").End(xlUp).Row
            countCDASheet.Range("o1", "u" & LastRow).ClearContents
        End If
        LastRow = SummarySheet.Cells(Cells.Rows.Count, "i").End(xlUp).Row
        SummarySheet.Range("b5", "ac" & LastRow + 1).ClearContents
    End If
            
            
    With SummarySheet
        
        LastRow = .Cells.Find(What:="*", After:=.Cells.Range("A1"), SearchDirection:=xlPrevious, LookIn:=xlFormulas, SearchOrder:=xlByRows).Row
        
        LastRow = Cells(Cells.Rows.Count, "b").End(xlUp).Row
        
    ' Create a new workbook and set a variable to the first sheet.

    ' Open the file dialog box and filter on Excel files, allowing multiple files
    ' to be selected.
        On Error Resume Next
        SelectedFiles = Application.GetOpenFilename( _
            filefilter:="Excel Files (*.xl*), *.xl*", MultiSelect:=True)
            
        If IsArray(SelectedFiles) = False Then Exit Sub
    ' NRow keeps track of where to insert new rows in the destination workbook.
        NRow = LastRow + 1
        
        
        answer = MsgBox("������ ���ϵ���" + ComboBox1.Value + "�� �����Ͻðڽ��ϱ�?", vbYesNo + vbQuestion, "�۾� ����")
        If answer = vbYes Then
        
            ' Loop through the list of returned file names
            For NFile = LBound(SelectedFiles) To UBound(SelectedFiles)
                '������� ������ ���� ���� �̸� ����
                FileName = SelectedFiles(NFile)
                Set WorkBk = Application.Workbooks.Open(FileName)
                countCDA = 0
                countSecure = 0
                countNonSecure = 0
                '�� ������ ��ũ�� ��Ʈ����
                Dest_LastRow = WorkBk.Sheets(1).Cells(Cells.Rows.Count, "b").End(xlUp).Row
                '��Ʈ�� ���� �� ������ ���
                '7��°�ٺ��� ��Ʈ�� �ִ� �ٱ��� for��
                For i = 7 To Dest_LastRow
                '���� Q���� ���� CDA�̸� SummarySheet�� NRow�׿� ���� / ȣ�⺰�� ����Ǿ������� CDAī��Ʈ
                '�پ� ��ũ�� ����
                    If WorkBk.Sheets(1).Range("q" & i).Value = "CDA" Then
                        SummarySheet.Range("B" & NRow, "AA" & NRow).Value = WorkBk.Sheets(1).Range("b" & i, "AA" & i).Value
                        NRow = NRow + 1
                        countCDA = countCDA + 1
                        If WorkBk.Sheets(1).Cells(i, 5) = "����" Then
                            countSecure = countSecure + 1
                        ElseIf WorkBk.Sheets(1).Cells(i, 5) = "�����" Then
                            countNonSecure = countNonSecure + 1
                        End If
                        
                    End If
                Next
                If b = "����5ȣ��" Then
                    Dest_LastRow = countCDASheet.Cells(Cells.Rows.Count, "i").End(xlUp).Row
                    countCDASheet.Cells(Dest_LastRow + 1, "i") = FileNameExtOf(FileName)
                    countCDASheet.Cells(Dest_LastRow + 1, "j") = countSecure
                    countCDASheet.Cells(Dest_LastRow + 1, "k") = countNonSecure
                    countCDASheet.Cells(Dest_LastRow + 1, "l") = countCDA - (countSecure + countNonSecure)
                    countCDASheet.Cells(Dest_LastRow + 1, "m") = countCDA
                    countCDASheet.Cells(Dest_LastRow + 1, "n") = b
                ElseIf b = "����6ȣ��" Then
                    Dest_LastRow = countCDASheet.Cells(Cells.Rows.Count, "p").End(xlUp).Row
                    countCDASheet.Cells(Dest_LastRow + 1, "p") = FileNameExtOf(FileName)
                    countCDASheet.Cells(Dest_LastRow + 1, "q") = countSecure
                    countCDASheet.Cells(Dest_LastRow + 1, "r") = countNonSecure
                    countCDASheet.Cells(Dest_LastRow + 1, "s") = countCDA - (countSecure + countNonSecure)
                    countCDASheet.Cells(Dest_LastRow + 1, "t") = countCDA
                    countCDASheet.Cells(Dest_LastRow + 1, "u") = b
                End If
                
                WorkBk.Close savechanges:=False
            Next NFile
    
            ' Call AutoFit on the destination sheet so that all data is readable.
            SummarySheet.Columns.AutoFit
        End If
    End With

End Sub

Public Function FileNameExtOf(ByVal s As String) As String
    FileNameExtOf = Mid$(s, InStrRev(s, "\") + 1)
End Function


Private Sub UserForm_Initialize()
    ComboBox1.AddItem "����5ȣ��"
    ComboBox1.AddItem "����6ȣ��"
    
    ComboBox2.AddItem "SR"
    ComboBox2.AddItem "ITS(ITS-5)"
    ComboBox2.AddItem "EP"
    ComboBox2.AddItem "Support"
    
    ComboBox3.AddItem "SSEP���"
    ComboBox3.AddItem "SSEP���/CS/CDA�� �ǿ���"
    ComboBox3.AddItem "CS/CDA ���ٰ�� ����"
    ComboBox3.AddItem "CS/CDA ����"
    ComboBox3.AddItem "�ý��� ��ȣ"
    
    ComboBox4.AddItem "EP"
    ComboBox4.AddItem "BOP"
    ComboBox4.AddItem "Indirect CDA"
    ComboBox4.AddItem "Direct CDA"
End Sub
