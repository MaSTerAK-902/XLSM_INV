Attribute VB_Name = "Module1"
Sub ���̓t�H�[��()
Attribute ���̓t�H�[��.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ���̓t�H�[�� Macro
'
    Worksheets("���͉��").Select
    Worksheets("���͉��").Activate
'
End Sub
Sub �i���ǉ�()
Attribute �i���ǉ�.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �i���ǉ� Macro
'
    Worksheets("�i���ǉ�").Select
    Worksheets("�i���ǉ�").Activate
'
End Sub
Sub �݌ɏ��()
Attribute �݌ɏ��.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �݌ɏ�� Macro
'
    Worksheets("�݌ɏ��").Select
    Worksheets("�݌ɏ��").Activate
'
End Sub
Sub ���׏�()
Attribute ���׏�.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ���׏� Macro
'
    Worksheets("���׏�").Select
    Worksheets("���׏�").Activate
'
End Sub
Sub �ۑ�()
Attribute �ۑ�.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �ۑ� Macro
'
    ThisWorkbook.Save
'
End Sub
Sub �I��()
Attribute �I��.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �I�� Macro
'
    ThisWorkbook.Close
'
End Sub
Sub �o��_���͎��s()
Attribute �o��_���͎��s.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �o�ɓ��͎��s Macro
'
'�ϐ��錾
Dim KeyName As String '�������[�h
Dim InDate As Date
Dim InStock As Long
Dim InReson As String
Dim RowNum As Long

Dim Result_Stock As Long
Dim Result_Stock_Sum As Long
Dim Result_Price As Long
Dim Result_Price_Sum As Long

Dim SearchRange As Range '�����͈͊i�[
Dim Result As Range '�����͈͂̔z��

Set ws1 = Worksheets("���͉��")
Set ws2 = Worksheets("�X�V����")
Set ws3 = Worksheets("�݌ɏ��")

'���͒l�̊i�[
KeyName = ws1.Range("B12").Value
InDate = ws1.Range("E12").Value
InStock = ws1.Range("G12").Value
Reason = ws1.Range("I12").Value



'�݌ɏ����X�V����

Set ResultRange = ws3.Range("A:A").Find(KeyName, LookAt:=xlWhole) '�ŏ��Ɉ�v����Range���擾

If ResultRange Is Nothing Then '�������ʂ𔻒�

    MsgBox "�������ʂȂ�"
    
    Exit Sub

Else
    If ws2.AutoFilterMode = True Then 'AutoFilter�̉���
        ws2.Range("A1").AutoFilter
    Else
    End If

    '�݌Ƀf�[�^�̎擾
    Result_Stock = ws3.Range("D" & ResultRange.Row)
    Result_Price = ws3.Range("C" & ResultRange.Row)
    
    '�݌ɂO�̏ꍇ
    If Result_Stock <= 0 Then
        MsgBox "�݌ɐ����O�ł��B"
        Exit Sub
    End If
    
    '�v�Z
    Result_Stock_Sum = Result_Stock - InStock '����
    Result_Price_Sum = Result_Stock_Sum * Result_Price '�݌ɋ��z
    
    '�u��
    ws3.Range("D" & ResultRange.Row) = Result_Stock_Sum
    ws3.Range("E" & ResultRange.Row) = Result_Price_Sum
    
End If

'�X�V�����ɏ���ǉ�����
If ws2.AutoFilterMode = True Then 'AutoFilter�̉���
    ws2.Range("A1").AutoFilter
Else

End If

RowNum = ws2.Cells(Rows.Count, "A").End(xlUp).Row + 1 '�s�̍ŉ��Ɉړ�

ws2.Range("A" & RowNum).Offset(0, 0) = KeyName
ws2.Range("A" & RowNum).Offset(0, 1) = InDate
ws2.Range("A" & RowNum).Offset(0, 2) = 0
ws2.Range("A" & RowNum).Offset(0, 3) = InStock
ws2.Range("A" & RowNum).Offset(0, 4) = Reason

'���͒l������
ws1.Range("G12") = 0

End Sub
Sub �i���ǉ�_���͎��s()
Attribute �i���ǉ�_���͎��s.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �i���ǉ�_���͎��s Macro
'
'�ϐ��錾
Dim KeyName As String '�������[�h
Dim CostPrice As Long
Dim SellPrice As Long
Dim RowNum As Long

Dim SearchRange As Range '�����͈͊i�[
Dim Result As Range '�����͈͂̔z��

Set ws1 = Worksheets("���͉��")
Set ws2 = Worksheets("�i���ǉ�")
Set ws3 = Worksheets("�݌ɏ��")

'���͒l�̊i�[
KeyName = ws2.Range("B6").Value
CostPrice = ws2.Range("E6").Value
SellPrice = ws2.Range("G6").Value

'�݌ɏ����X�V����
If ws3.AutoFilterMode = True Then 'AutoFilter�̉���
    ws3.Range("A1").AutoFilter
End If

Set ResultRange = ws3.Range("A:A").Find(KeyName, LookAt:=xlWhole) '�ŏ��Ɉ�v����Range���擾

If ResultRange Is Nothing Then '�������ʂ𔻒�
    
    RowNum = ws3.Cells(Rows.Count, "A").End(xlUp).Row + 1 '�s�̍ŉ��Ɉړ�
    
    ws3.Range("A" & RowNum).Offset(0, 0) = KeyName
    ws3.Range("A" & RowNum).Offset(0, 1) = SellPrice
    ws3.Range("A" & RowNum).Offset(0, 2) = CostPrice
    ws3.Range("A" & RowNum).Offset(0, 3) = 0
    ws3.Range("A" & RowNum).Offset(0, 4) = 0
    
    MsgBox "�ǉ����܂����B"
    
    Exit Sub

Else
    MsgBox "�i���̏d��������܂����B"
    Exit Sub
End If

'
End Sub
Sub ����_���͎��s()
'
' ���ɓ��͎��s Macro
'
'�ϐ��錾
Dim KeyName As String '�������[�h
Dim InDate As Date
Dim InStock As Long
Dim InReson As String
Dim RowNum As Long

Dim Result_Stock As Long
Dim Result_Stock_Sum As Long
Dim Result_Price As Long
Dim Result_Price_Sum As Long

Dim SearchRange As Range '�����͈͊i�[
Dim Result As Range '�����͈͂̔z��

Set ws1 = Worksheets("���͉��")
Set ws2 = Worksheets("�X�V����")
Set ws3 = Worksheets("�݌ɏ��")

KeyName = ws1.Range("B8").Value
InDate = ws1.Range("E8").Value
InStock = ws1.Range("G8").Value
Reason = ws1.Range("I8").Value

'�݌ɏ����X�V����

Set ResultRange = ws3.Range("A:A").Find(KeyName, LookAt:=xlWhole) '�ŏ��Ɉ�v����Range���擾

If ResultRange Is Nothing Then '�������ʂ𔻒�

    MsgBox "�������ʂȂ�"
    
    Exit Sub

Else
    If ws3.AutoFilterMode = True Then 'AutoFilter�̉���
        ws3.Range("A1").AutoFilter
    Else
    End If
    
    '�݌Ƀf�[�^�̎擾
    Result_Stock = ws3.Range("D" & ResultRange.Row)
    Result_Price = ws3.Range("C" & ResultRange.Row)
    
    '�v�Z
    Result_Stock_Sum = Result_Stock + InStock '����
    Result_Price_Sum = Result_Stock_Sum * Result_Price '�݌ɋ��z
    
    '�u��
    ws3.Range("D" & ResultRange.Row) = Result_Stock_Sum
    ws3.Range("E" & ResultRange.Row) = Result_Price_Sum
    
End If

'�X�V�����ɏ���ǉ�����
If ws2.AutoFilterMode = True Then 'AutoFilter�̉���
    ws2.Range("A1").AutoFilter
End If

RowNum = ws2.Cells(Rows.Count, "A").End(xlUp).Row + 1 '�s�̍ŉ��Ɉړ�

ws2.Range("A" & RowNum).Offset(0, 0) = KeyName
ws2.Range("A" & RowNum).Offset(0, 1) = InDate
ws2.Range("A" & RowNum).Offset(0, 2) = InStock
ws2.Range("A" & RowNum).Offset(0, 3) = 0
ws2.Range("A" & RowNum).Offset(0, 4) = Reason

'���͒l������
ws1.Range("G8") = 0

End Sub
Sub ���ׂ̍쐬()
Attribute ���ׂ̍쐬.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ���ׂ̍쐬 Macro
'
'�ϐ��錾
Dim Start_date As Date
Dim Last_date As Date

Dim KeyName As String
Dim Sell_price As Long
Dim Cost_price As Long
Dim Stock_Now As Long

Dim InStock As Long '����
Dim SellStock As Long '�̔���
Dim ReturnStock As Long '�ԋp
Dim LostStock As Long '����
Dim OutStock As Long '���i
Dim BeforeStock As Long '�O�݌�
Dim Stock_Status As String '���o�ɗ��R

Dim RangeTitle As Range
Dim RangeDate As Range
Dim RangeKeyName As Range
Dim TodayDate As Date

Dim LastRow As Long
Dim Filter_Row As Long
Dim PrintRow As Long
Dim SellSum As Long

Dim i As Long
Dim n As Long

Set ws1 = Worksheets("���׏�")
Set ws2 = Worksheets("�݌ɏ��")
Set ws3 = Worksheets("�X�V����")
Set ws4 = Worksheets("����y�[�W")

'���׏�����y�[�W�̏�����
ws4.Cells.ClearContents
ws4.Range("A:I").Borders.LineStyle = xlLineStyleNone

'���͒l�̊i�[
Start_date = ws1.Range("B8").Value
Last_date = ws1.Range("E8").Value

'���[�v����
LastRow = ws2.Cells(Rows.Count, 1).End(xlUp).Row 'A��̍ŏI�s���擾

For i = 2 To LastRow '�w�b�_�[�͊O��

    'ws2�̃I�[�g�t�B���^�[������
    If ws2.AutoFilterMode = True Then 'AutoFilter�̉���
        ws2.Range("A1").AutoFilter
    End If

    '�݌ɏ��̎擾
    KeyName = ws2.Cells(i, 1)
    Sell_price = ws2.Cells(i, 2)
    Cost_price = ws2.Cells(i, 3)
    Stock_Now = ws2.Cells(i, 4)
    
    '�����l�ݒ�
    InStock = 0
    SellStock = 0
    ReturnStock = 0
    LostStock = 0
    OutStock = 0
    BeforeStock = 0
    
    'ws3�̃I�[�g�t�B���^�[������
    If ws3.AutoFilterMode = True Then 'AutoFilter�̉���
        ws3.Range("A1").AutoFilter
    End If
    
    '�X�V�����̏���(�I�[�g�t�B���^�[)
    ws3.Range("A1").AutoFilter Field:=1, _
    Criteria1:=KeyName
    ws3.Range("A1").AutoFilter Field:=2, _
    Criteria1:=">=" & Start_date, _
    Operator:=xlAnd, _
    Criteria2:="<=" & Last_date
    
    '�w�b�_�[�������s���̐ݒ�
    n = 2
    
    '�t�B���^�[���ʂ̏W�v
    Do While ws3.Cells(n, 1) <> ""
        
        'Cells(n, 1)��EntireRow.Hidden��False�Ȃ���s
        If ws3.Cells(n, 1).EntireRow.Hidden = False Then
        
            '�݌ɏ�Ԃ̎擾
            Stock_Status = ws3.Cells(n, 5).Value
            
            '�ꍇ����
            Select Case Stock_Status
            
                Case "����"
                
                    InStock = InStock + ws3.Cells(n, 3).Value
                    
                Case "�̔���"
                
                    SellStock = SellStock + ws3.Cells(n, 4).Value
                    
                Case "�ԋp"
                
                    ReturnStock = ReturnStock + ws3.Cells(n, 4).Value
                    
                Case "����"
                
                    LostStock = LostStock + ws3.Cells(n, 4).Value
                
                Case "���i"
                
                    OutStock = OutStock + ws3.Cells(n, 4).Value
        
            End Select
        
        End If
        
        n = n + 1
    Loop
    
    '�W�v���ʂ̍쐬(�O��݌�)
        
    BeforeStock = Stock_Now + SellStock + ReturnStock + LostStock + OutStock - InStock
    
    '���׏��̍쐬
    
    ws4.Cells(i + 3, 1) = KeyName
    ws4.Cells(i + 3, 1).HorizontalAlignment = xlLeft
    ws4.Cells(i + 3, 1).WrapText = True
    ws4.Cells(i + 3, 2) = Sell_price
    ws4.Cells(i + 3, 3) = Cost_price
    ws4.Cells(i + 3, 4) = BeforeStock
    ws4.Cells(i + 3, 5) = InStock
    ws4.Cells(i + 3, 6) = SellStock
    ws4.Cells(i + 3, 7) = Cost_price * SellStock
    ws4.Cells(i + 3, 8) = ReturnStock
    ws4.Cells(i + 3, 9) = Stock_Now
    
Next i

'ws2�̃I�[�g�t�B���^�[������
If ws2.AutoFilterMode = True Then 'AutoFilter�̉���
    ws2.Range("A1").AutoFilter
End If

'ws3�̃I�[�g�t�B���^�[������
If ws3.AutoFilterMode = True Then 'AutoFilter�̉���
    ws3.Range("A1").AutoFilter
End If
    
'���׏��g�g�݂̍쐬
Set RangeTitle = ws4.Range("A1:I2")
Set RangeDate = ws4.Range("G3:I3")

TodayDate = Date

RangeTitle.MergeCells = True
RangeDate.MergeCells = True

ws4.Range("A1").HorizontalAlignment = xlCenter
ws4.Range("A1").Value = "�^�C�g��"
ws4.Range("A1").Font.Size = 18

ws4.Range("G3").HorizontalAlignment = xlCenter
ws4.Range("G3") = "���t�F" & TodayDate

ws4.Range("A4").HorizontalAlignment = xlCenter
ws4.Range("A4").Value = "�i��"

ws4.Range("B4").HorizontalAlignment = xlCenter
ws4.Range("B4").Value = "���l"

ws4.Range("C4").HorizontalAlignment = xlCenter
ws4.Range("C4").Value = "�d���l"

ws4.Range("D4").HorizontalAlignment = xlCenter
ws4.Range("D4").Value = "�O��݌�"

ws4.Range("E4").HorizontalAlignment = xlCenter
ws4.Range("E4").Value = "�V�K��"

ws4.Range("F4").HorizontalAlignment = xlCenter
ws4.Range("F4").Value = "�̔���"

ws4.Range("G4").HorizontalAlignment = xlCenter
ws4.Range("G4").Value = "���v"

ws4.Range("H4").HorizontalAlignment = xlCenter
ws4.Range("H4").Value = "�ԋp"

ws4.Range("I4").HorizontalAlignment = xlCenter
ws4.Range("I4").Value = "���݌�"

'A��̍ŏI�s���擾
PrintRow = ws4.Cells(Rows.Count, 1).End(xlUp).Row

ws4.Range("A4:I4").Borders(xlEdgeTop).Weight = xlMedium
ws4.Range("A4:I4").Borders(xlEdgeLeft).Weight = xlMedium
ws4.Range("A4:I4").Borders(xlEdgeBottom).Weight = xlMedium
ws4.Range("A4:I4").Borders(xlEdgeRight).Weight = xlMedium
ws4.Range("A4:I4").Borders(xlInsideVertical).Weight = xlMedium

ws4.Range(ws4.Cells(5, 1), ws4.Cells(PrintRow, 9)).Borders(xlEdgeTop).Weight = xlMedium
ws4.Range(ws4.Cells(5, 1), ws4.Cells(PrintRow, 9)).Borders(xlEdgeLeft).Weight = xlMedium
ws4.Range(ws4.Cells(5, 1), ws4.Cells(PrintRow, 9)).Borders(xlEdgeBottom).Weight = xlMedium
ws4.Range(ws4.Cells(5, 1), ws4.Cells(PrintRow, 9)).Borders(xlEdgeRight).Weight = xlMedium
ws4.Range(ws4.Cells(5, 1), ws4.Cells(PrintRow, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous
ws4.Range(ws4.Cells(5, 1), ws4.Cells(PrintRow, 9)).Borders(xlInsideHorizontal).LineStyle = xlContinuous

ws4.Range(ws4.Cells(5, 1), ws4.Cells(PrintRow, 1)).Borders(xlEdgeRight).Weight = xlMedium

'���v�̍��v��ǉ�����B
ws4.Range(ws4.Cells(PrintRow + 1, 7), ws4.Cells(PrintRow + 1, 9)).MergeCells = True
SellSum = WorksheetFunction.Sum(ws4.Range(ws4.Cells(5, 7), ws4.Cells(PrintRow, 7)))
ws4.Cells(PrintRow + 1, 7) = SellSum
ws4.Cells(PrintRow + 1, 7).HorizontalAlignment = xlLeft
ws4.Range(ws4.Cells(PrintRow + 1, 7), ws4.Cells(PrintRow + 1, 9)).Borders(xlEdgeTop).Weight = xlMedium
ws4.Range(ws4.Cells(PrintRow + 1, 7), ws4.Cells(PrintRow + 1, 9)).Borders(xlEdgeLeft).Weight = xlMedium
ws4.Range(ws4.Cells(PrintRow + 1, 7), ws4.Cells(PrintRow + 1, 9)).Borders(xlEdgeBottom).Weight = xlMedium
ws4.Range(ws4.Cells(PrintRow + 1, 7), ws4.Cells(PrintRow + 1, 9)).Borders(xlEdgeRight).Weight = xlMedium

ws4.Cells(PrintRow + 1, 6) = "���v"
ws4.Cells(PrintRow + 1, 6).HorizontalAlignment = xlCenter
ws4.Cells(PrintRow + 1, 6).Borders(xlEdgeTop).Weight = xlMedium
ws4.Cells(PrintRow + 1, 6).Borders(xlEdgeLeft).Weight = xlMedium
ws4.Cells(PrintRow + 1, 6).Borders(xlEdgeBottom).Weight = xlMedium
ws4.Cells(PrintRow + 1, 6).Borders(xlEdgeRight).Weight = xlMedium

ws4.Activate

End Sub

Sub ����()
'
' ���� Macro
'
    Sheets(5).Select
    Sheets(5).Activate
'
End Sub
Sub �`�F�b�N�{�b�N�X�쐬()

Dim StartX As Single
Dim StartY As Single
Dim EndX As Single
Dim EndY As Single
Dim i As Long
Dim LastRow As Long

    'A��̍ŏI�s���擾
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'B��Ƀ`�F�b�N�{�b�N�X���쐬����
    For i = 1 To LastRow

        With Cells(i, 6)

            '�Z���̍��[
            StartX = .Left

            '�Z���̏�[
            StartY = .Top

            '�Z���̉���
            EndX = .Offset(0, 1).Left - .Left

            '�Z���̍���
            EndY = .Height

            '�`�F�b�N�{�b�N�X���
            ActiveSheet.CheckBoxes.Add(StartX, StartY, EndX, EndY).Select

            '�`�F�b�N�{�b�N�X�̃e�L�X�g���w��
            Selection.Text = ""

            '�Z���ɍ��킹�Ĉړ���T�C�Y��ύX����
            Selection.Placement = xlMoveAndSize

        End With

    Next i

End Sub

