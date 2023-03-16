Attribute VB_Name = "budget_enter"
Const MAIN_SHEET_NAME As String = "BUDGET"
Const DELIVERABLE_SHEET_NAME As String = "QUERY"
Const DEPT_MASTER_SHEET_NAME As String = "����}�X�^"
Const ASSET_MASTER_SHEET_NAME As String = "�Ȗڃ}�X�^"
Const ASEETS_TOTAL As Integer = 58

Private Sub BudgetEnter()

'BUDGET�V�[�g���w��
If Not ExistsSheet(MAIN_SHEET_NAME) Then
    MsgBox (MAIN_SHEET_NAME & "�V�[�g��������܂���B")
    End
Else
    Dim main As Worksheet: Set main = Worksheets(MAIN_SHEET_NAME)
End If

'�N�x�̊m�F
Dim year As String: year = Left(main.Cells(2, 2), 4)
Dim next_year As String: next_year = CStr(Val(year) + 1)
Dim rc As Integer
rc = MsgBox(year & "�N�x�̏������s���܂����H", vbYesNo + vbQuestion, "�m�F")
If Not rc = vbYes Then
    MsgBox "�����𒆒f���܂���"
    End
End If

'�L�[�ƂȂ�̂�O��(15)��AF��(32)
Dim keys() As Variant
keys = Array(15, 32)

For Each Key In keys

    For i = 3 To main.Cells(Rows.Count, Key).End(xlUp).row
        '����̈�̒T��
        '�e�Z�������喼�ɂȂ��Ă��邩����(�G���[�l�͔�΂�)
        dept_name = main.Cells(i, Key)
        If IsError(dept_name) Then
            GoTo CONTINUE:
        End If
        Dim dept_code As Integer: dept_code = ExistsDept(dept_name)
        If dept_code <> 0 Then
            '2��->3���܂ŋt����
            For j = Key To Key - 11 Step -1
                '2���Z�������Z���ɂ���
                Dim base_cell As Range: Set base_cell = main.Cells(i + 1, j)
                Dim y As String
                Dim m As String:  m = TransMonth(base_cell.Value)
                If m = "01" Or m = "02" Then
                    y = next_year
                Else
                    y = year
                End If

                'QUERY�t�H���_�ɏo��
                For row = 1 To ASEETS_TOTAL
                    Dim budget As Double: budget = base_cell.Offset(row, 0).Value * 10000
                    QueryBuilder dept_code, y & m, row + 9, year, Round(budget, 0)
                Next
            Next j
        End If
CONTINUE:
    Next
Next Key

'�����ʒm
MsgBox ("����I�����܂���")

End Sub

'�N�G���r���_�[
Public Function QueryBuilder(ByVal dept_code As String, ByVal ym As String, ByVal asset_code As Integer, ByVal year As String, ByVal budget As Double)
    Dim asset_master As Worksheet: Set asset_master = Worksheets(ASSET_MASTER_SHEET_NAME)
    Dim query As Worksheet: Set query = Worksheets(DELIVERABLE_SHEET_NAME)
    
    'Query�V�[�g�̍ŏI�s���擾
    Dim edit_position As Integer: edit_position = query.Cells(Rows.Count, 1).End(xlUp).row + 1
    Dim target As Range: Set target = query.Cells(edit_position, 1)
    
    'INSERT����1000���R�[�h���Ƃɋ�؂��Ă�����
    If edit_position Mod 1000 = 0 Then
        target.Offset(1, 0) = "SELECT * FROM DUAL;"
        target.Offset(2, 0) = "INSERT ALL"
        target = target.Offset(4, 0)
    End If
    
    target.Formula = "INTO M_YOSAN(M_YOSAN.YO_JGYCD,M_YOSAN.YO_BMNCD,M_YOSAN.YO_YM, M_YOSAN.YO_KCKBN, M_YOSAN.YO_KMKCD, M_YOSAN.YO_KMKNM, M_YOSAN.YO_NENDO, M_YOSAN.YO_YOSAN ) VALUES ("
    
    target.Offset(0, 1).Formula = "=VLOOKUP(" & target.Offset(0, 2).Address & "," & DEPT_MASTER_SHEET_NAME & "!B:C,2,)&"","""
    
    target.Offset(0, 2).Formula = dept_code
    
    target.Offset(0, 3).Formula = "," & ym & ",0,"
    
    target.Offset(0, 4).Formula = asset_code
    
    target.Offset(0, 5).Formula = ",'"

    target.Offset(0, 6).Formula = "=VLOOKUP(" & target.Offset(0, 4).Address & "," & ASSET_MASTER_SHEET_NAME & "!A:B,2,)"

    target.Offset(0, 7).Formula = "''," & year & "," & budget & ")"
    
End Function

' Sheets �Ɏw�肵�����O�̃V�[�g�����݂��邩����
Public Function ExistsSheet(ByVal bookName As String)
    Dim ws As Variant
    For Each ws In Sheets
        If LCase(ws.Name) = LCase(bookName) Then
            ExistsSheet = True ' ���݂���
            Exit Function
        End If
    Next

    ExistsSheet = False
End Function

' �w�肵������̑��݃`�F�b�N
Public Function ExistsDept(ByVal deptName As String) As Integer
    Dim dept_master As Worksheet: Set dept_master = Worksheets(DEPT_MASTER_SHEET_NAME)
    For i = 2 To dept_master.Cells(Rows.Count, 1).End(xlUp).row
        
        If dept_master.Cells(i, 1) = deptName Then
            ExistsDept = dept_master.Cells(i, 2)
            Exit Function
        End If
    Next
    
    ExistsDept = 0
End Function

'���x�t�H�[�}�b�^
Public Function TransMonth(ByVal month As String) As String
    Dim monthIndex As String

    Select Case month
        Case "�Q��"
            monthIndex = "02"
        Case "�P��"
            monthIndex = "01"
        Case "�P�Q��"
            monthIndex = "12"
        Case "�P�P��"
            monthIndex = "11"
        Case "�P�O��"
            monthIndex = "10"
        Case "�X��"
            monthIndex = "09"
        Case "�W��"
            monthIndex = "08"
        Case "�V��"
            monthIndex = "07"
        Case "�U��"
            monthIndex = "06"
        Case "�T��"
            monthIndex = "05"
        Case "�S��"
            monthIndex = "04"
        Case "�R��"
            monthIndex = "03"
        Case Else
            MsgBox ("���x������" & month & "�͕s�K�؂ł�")
            End
    End Select

    TransMonth = monthIndex
End Function
