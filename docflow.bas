Attribute VB_Name = "Module1"
Sub docflow()
Attribute docflow.VB_ProcData.VB_Invoke_Func = "q\n14"
Dim value As Variant
Dim val2 As Variant
Dim sFilesPath As String

    '�������� �������
    Sheets("��� �������� ���������������").Select
    Range("A4").Select
    Selection.EntireRow.Insert
    
        '������������ �����
        Sheets("��� �������� ���������������").Select
        Range("B4").Select
        value = InputBox("'��'=��������� ������� ��������;" & vbCrLf & "'��'=��������� ����������� ��������," & vbCrLf & " ** = �����", "������������ ����� ")
        If value = "**" Then Exit Sub
        If value = "��" Then
            ActiveCell.value = "��������� ����������� ��������"
        ElseIf value = "��" Then
            ActiveCell.value = "��������� ������� ��������"
        Else
            ActiveCell.value = value
        End If
        
        
    '��������� �����
    Sheets("��� �������� ���������������").Select
    Range("C4").Select
    value = InputBox("� �����������" & vbCrLf & "** = �����", "���������")
    If value = "**" Then Exit Sub
    ActiveCell.value = value

    With Application.FileDialog(msoFileDialogFilePicker)
        
       If .Show = False Then Exit Sub
       .Filters.Add "All Files", "*.*"
       sFilesPath = .SelectedItems(1)
   End With
      ActiveSheet.Hyperlinks.Add Anchor:=Range("C4"), Address:=sFilesPath
    
        '��������� ����
        Sheets("��� �������� ���������������").Select
        Range("D4").Select
        value = InputBox("���� �����������" & vbCrLf & " 0=�������" & vbCrLf & "�����= ����� ����� ������, ������: ��� ����� 19 ����� ��������� 19.04.2018" & vbCrLf & "** = �����", "���������")
        If value = "**" Then Exit Sub
            '0=�������
        If value = 0 Then value = Date
            '���� ����� = ���� ����� ������
        If IsNumeric(value) Then value = DateSerial(Year(Date), Month(Date), value)
        ActiveCell.value = value
    
    '� ����� �� �����
    Sheets("��� �������� ���������������").Select
    Range("E4").Select
    value = InputBox("� �����������" & vbCrLf & "** = �����", "� ����� ��")
    If value = "**" Then Exit Sub
    ActiveCell.value = value
    
        '� ����� �� ����
        Sheets("��� �������� ���������������").Select
        Range("F4").Select
        value = InputBox("���� �����������" & vbCrLf & "0=�������" & vbCrLf & "�����= ����� ����� ������" & vbCrLf & "������: ��� ����� 19 ����� ��������� 19.04.2018" & vbCrLf & "** = �����", "� ����� ��")
        If value = "**" Then Exit Sub
            '0=�������
        If value = 0 Then value = Date
            '���� ����� = ���� ����� ������
        If IsNumeric(value) Then value = DateSerial(Year(Date), Month(Date), value)
        ActiveCell.value = value

    '�������� �����
    Sheets("��� �������� ���������������").Select
    Range("G4").Select
    value = InputBox("� �����������" & vbCrLf & "** = �����", "��������")
    If value = "**" Then Exit Sub
    ActiveCell.value = value
    
        '�������� ����
        Sheets("��� �������� ���������������").Select
        Range("H4").Select
        value = InputBox("���� �����������" & vbCrLf & "0=�������" & vbCrLf & "�����= ����� ����� ������, ������: ��� ����� 19 ����� ��������� 19.04.2018" & vbCrLf & "** = �����", "��������")
        If value = "**" Then Exit Sub
            '0=�������
        If value = 0 Then value = Date
            '���� ����� = ���� ����� ������
        If IsNumeric(value) Then value = DateSerial(Year(Date), Month(Date), value)
        ActiveCell.value = value

    '���� (�������)(� ��. ������)
    Sheets("��� �������� ���������������").Select
    Range("I4").Select
    value = InputBox("������� � ��. ������" & vbCrLf & "�� = �������" & vbCrLf & "�� = ���������" & vbCrLf & "��+ = ��������� � �������� ���" & vbCrLf & "** = �����", "����")
    Select Case value
        Case "**"
            Exit Sub
        Case "��"
            value = "�.�. ��������"
        Case "��"
            value = "�.�. ����������"
        Case "��+"
            val2 = InputBox("����������/..." & vbCrLf & "�� = �����������")
            If val2 = "��" Then
                value = "�.�.����������/ �.�. �����������"
            Else
                value = "�.�.����������/ " + val2
            End If
        End Select
    ActiveCell.value = value

        '���� (�����������)(� ��. ������)
        Sheets("��� �������� ���������������").Select
        Range("J4").Select
        value = InputBox("�����������" & vbCrLf & "��� = �������������" & vbCrLf & "���+ = ������������� + ��� ��� ��" & vbCrLf & "** = �����", "����")
        Select Case value
            Case "**"
                Exit Sub
            Case "���"
                value = "�������������"
            Case "���+"
                val2 = InputBox("�������������/..." & vbCrLf & "�� = �.�������� �" & vbCrLf & "��� = ��������� �" & vbCrLf & _
                "��� = ������ �" & vbCrLf & "��� = ��������� �" & vbCrLf & "��� = �������� �" & vbCrLf & "��� = ������� �" & vbCrLf & "�� = ���������� �" & vbCrLf & _
                "�� = ������ �" & vbCrLf & "��� = ����������� �" & vbCrLf & "** = �����")
                Select Case val2
                Case "**"
                    Exit Sub
                Case "��"
                    val2 = "������������� ������"
                Case "���"
                    val2 = "�����-������������� ������"
                Case "���"
                    val2 = "���������� ������"
                Case "���"
                    val2 = "�� ������������"
                Case "��"
                    val2 = "��� ������� ������"
                Case "���"
                    val2 = "�������������� ������"
                Case "���"
                    val2 = "���������� ������"
                Case "���"
                    val2 = "����������� ������"
                Case "��"
                    val2 = "��������� ������"
                Case "���"
                    val2 = "������������� ������"
                Case "��"
                    val2 = "�������������� ������"
                End Select
                value = "�������������/ " + val2
            End Select
        ActiveCell.value = value
    
    '�� ���� - �����������
    Sheets("��� �������� ���������������").Select
    Range("L4").Select
    value = InputBox("�� ���� �����������" & vbCrLf & "�� = �.�������� �" & vbCrLf & _
        "��� = ��������� �" & vbCrLf & "��� = ������ �" & vbCrLf & "��� = ��������� �" & vbCrLf & _
        "��� = �������� �" & vbCrLf & "��� = ������� �" & vbCrLf & "�� = ���������� �" & vbCrLf & _
        "�� = ������ �" & vbCrLf & "��� = ����������� �" & vbCrLf & "��� = �������������" & vbCrLf & _
        "--------------" & vbCrLf & "�� = ������� ������" & vbCrLf & "��� = ������������" & vbCrLf & "** = �����", "�� ����")
    Select Case value
        Case "**"
            Exit Sub
        Case "��"
            value = "������������� ������"
        Case "���"
            value = "�����-������������� ������"
        Case "���"
            value = "���������� ������"
        Case "���"
            value = "�� ������������"
        Case "��"
            value = "��� ������� ������"
        Case "���"
            value = "�������������� ������"
        Case "���"
            value = "���������� ������"
        Case "���"
            value = "����������� ������"
        Case "��"
            value = "��������� ������"
        Case "���"
            value = "������������� ������"
        Case "��"
            value = "�������������� ������"
        End Select
    ActiveCell.value = value
    val2 = value
    
    '�� ���� - ���
    Sheets("��� �������� ���������������").Select
    Range("K4").Select
    Select Case val2
        Case "������������� ������"
            value = "�.�. �����"
        Case "�����-������������� ������"
            value = "�.�. �����"
        Case "���������� ������"
            value = "�.�. ��������"
        Case "�� ������������"
            value = "�.�. ��������"
        Case "��� ������� ������"
            value = "�.�. ��������"
        Case "�������������� ������"
            value = "�.�. ����������� "
        Case "���������� ������"
            value = "�.�. �����"
        Case "����������� ������"
            value = "�.�. �������"
        Case "��������� ������"
            value = "�.�. �������"
        Case "������������� ������"
            value = "�.�. �����"
        Case "�������������� ������"
            value = "�.�. �������"
        Case Else
            value = InputBox("��� �� ����" & vbCrLf & "** = �����", "�� ����")
            If value = "**" Then Exit Sub
    End Select
    ActiveCell.value = value

    '�������������
    Sheets("��� �������� ���������������").Select
    Range("N4").Select
    value = InputBox("�����������" & vbCrLf & "1 = �.�. ��������" & vbCrLf & "2 = ��������/����������" & vbCrLf & "3 = �.�. �������" & vbCrLf & "** = �����", "�����������")
    Select Case value
        Case "**"
            Exit Sub
        Case 1
            value = "�.�. ��������"
        Case 2
            value = "�.�. ��������/ �.�. ����������"
        Case 3
            value = "�.�. �������"
        End Select
    ActiveCell.value = value
    
    
End Sub
