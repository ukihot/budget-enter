Attribute VB_Name = "budget_enter"
Const MAIN_SHEET_NAME As String = "BUDGET"
Const DELIVERABLE_SHEET_NAME As String = "QUERY"
Const DEPT_MASTER_SHEET_NAME As String = "部門マスタ"
Const ASSET_MASTER_SHEET_NAME As String = "科目マスタ"
Const ASEETS_TOTAL As Integer = 58

Private Sub BudgetEnter()

'BUDGETシートを指定
If Not ExistsSheet(MAIN_SHEET_NAME) Then
    MsgBox (MAIN_SHEET_NAME & "シートが見つかりません。")
    End
Else
    Dim main As Worksheet: Set main = Worksheets(MAIN_SHEET_NAME)
End If

'年度の確認
Dim year As String: year = Left(main.Cells(2, 2), 4)
Dim next_year As String: next_year = CStr(Val(year) + 1)
Dim rc As Integer
rc = MsgBox(year & "年度の処理を行いますか？", vbYesNo + vbQuestion, "確認")
If Not rc = vbYes Then
    MsgBox "処理を中断しました"
    End
End If

'キーとなるのはO列(15)とAF列(32)
Dim keys() As Variant
keys = Array(15, 32)

For Each Key In keys

    For i = 3 To main.Cells(Rows.Count, Key).End(xlUp).row
        '部門領域の探索
        '各セルが部門名になっているか走査(エラー値は飛ばす)
        dept_name = main.Cells(i, Key)
        If IsError(dept_name) Then
            GoTo CONTINUE:
        End If
        Dim dept_code As Integer: dept_code = ExistsDept(dept_name)
        If dept_code <> 0 Then
            '2月->3月まで逆走査
            For j = Key To Key - 11 Step -1
                '2月セルを基底セルにする
                Dim base_cell As Range: Set base_cell = main.Cells(i + 1, j)
                Dim y As String
                Dim m As String:  m = TransMonth(base_cell.Value)
                If m = "01" Or m = "02" Then
                    y = next_year
                Else
                    y = year
                End If

                'QUERYフォルダに出力
                For row = 1 To ASEETS_TOTAL
                    Dim budget As Double: budget = base_cell.Offset(row, 0).Value * 10000
                    QueryBuilder dept_code, y & m, row + 9, year, Round(budget, 0)
                Next
            Next j
        End If
CONTINUE:
    Next
Next Key

'完了通知
MsgBox ("正常終了しました")

End Sub

'クエリビルダー
Public Function QueryBuilder(ByVal dept_code As String, ByVal ym As String, ByVal asset_code As Integer, ByVal year As String, ByVal budget As Double)
    Dim asset_master As Worksheet: Set asset_master = Worksheets(ASSET_MASTER_SHEET_NAME)
    Dim query As Worksheet: Set query = Worksheets(DELIVERABLE_SHEET_NAME)
    
    'Queryシートの最終行を取得
    Dim edit_position As Integer: edit_position = query.Cells(Rows.Count, 1).End(xlUp).row + 1
    Dim target As Range: Set target = query.Cells(edit_position, 1)
    
    'INSERT文は1000レコードごとに区切っておこう
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

' Sheets に指定した名前のシートが存在するか判定
Public Function ExistsSheet(ByVal bookName As String)
    Dim ws As Variant
    For Each ws In Sheets
        If LCase(ws.Name) = LCase(bookName) Then
            ExistsSheet = True ' 存在する
            Exit Function
        End If
    Next

    ExistsSheet = False
End Function

' 指定した部門の存在チェック
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

'月度フォーマッタ
Public Function TransMonth(ByVal month As String) As String
    Dim monthIndex As String

    Select Case month
        Case "２月"
            monthIndex = "02"
        Case "１月"
            monthIndex = "01"
        Case "１２月"
            monthIndex = "12"
        Case "１１月"
            monthIndex = "11"
        Case "１０月"
            monthIndex = "10"
        Case "９月"
            monthIndex = "09"
        Case "８月"
            monthIndex = "08"
        Case "７月"
            monthIndex = "07"
        Case "６月"
            monthIndex = "06"
        Case "５月"
            monthIndex = "05"
        Case "４月"
            monthIndex = "04"
        Case "３月"
            monthIndex = "03"
        Case Else
            MsgBox ("月度文字に" & month & "は不適切です")
            End
    End Select

    TransMonth = monthIndex
End Function
