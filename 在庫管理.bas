Attribute VB_Name = "Module1"
Sub 入力フォーム()
Attribute 入力フォーム.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 入力フォーム Macro
'
    Worksheets("入力画面").Select
    Worksheets("入力画面").Activate
'
End Sub
Sub 品名追加()
Attribute 品名追加.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 品名追加 Macro
'
    Worksheets("品名追加").Select
    Worksheets("品名追加").Activate
'
End Sub
Sub 在庫情報()
Attribute 在庫情報.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 在庫情報 Macro
'
    Worksheets("在庫情報").Select
    Worksheets("在庫情報").Activate
'
End Sub
Sub 明細書()
Attribute 明細書.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 明細書 Macro
'
    Worksheets("明細書").Select
    Worksheets("明細書").Activate
'
End Sub
Sub 保存()
Attribute 保存.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 保存 Macro
'
    ThisWorkbook.Save
'
End Sub
Sub 終了()
Attribute 終了.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 終了 Macro
'
    ThisWorkbook.Close
'
End Sub
Sub 出庫_入力実行()
Attribute 出庫_入力実行.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 出庫入力実行 Macro
'
'変数宣言
Dim KeyName As String '検索ワード
Dim InDate As Date
Dim InStock As Long
Dim InReson As String
Dim RowNum As Long

Dim Result_Stock As Long
Dim Result_Stock_Sum As Long
Dim Result_Price As Long
Dim Result_Price_Sum As Long

Dim SearchRange As Range '検索範囲格納
Dim Result As Range '検索範囲の配列

Set ws1 = Worksheets("入力画面")
Set ws2 = Worksheets("更新履歴")
Set ws3 = Worksheets("在庫情報")

'入力値の格納
KeyName = ws1.Range("B12").Value
InDate = ws1.Range("E12").Value
InStock = ws1.Range("G12").Value
Reason = ws1.Range("I12").Value



'在庫情報を更新する

Set ResultRange = ws3.Range("A:A").Find(KeyName, LookAt:=xlWhole) '最初に一致したRangeを取得

If ResultRange Is Nothing Then '検索結果を判定

    MsgBox "検索結果なし"
    
    Exit Sub

Else
    If ws2.AutoFilterMode = True Then 'AutoFilterの解除
        ws2.Range("A1").AutoFilter
    Else
    End If

    '在庫データの取得
    Result_Stock = ws3.Range("D" & ResultRange.Row)
    Result_Price = ws3.Range("C" & ResultRange.Row)
    
    '在庫０の場合
    If Result_Stock <= 0 Then
        MsgBox "在庫数が０です。"
        Exit Sub
    End If
    
    '計算
    Result_Stock_Sum = Result_Stock - InStock '結果
    Result_Price_Sum = Result_Stock_Sum * Result_Price '在庫金額
    
    '置換
    ws3.Range("D" & ResultRange.Row) = Result_Stock_Sum
    ws3.Range("E" & ResultRange.Row) = Result_Price_Sum
    
End If

'更新履歴に情報を追加する
If ws2.AutoFilterMode = True Then 'AutoFilterの解除
    ws2.Range("A1").AutoFilter
Else

End If

RowNum = ws2.Cells(Rows.Count, "A").End(xlUp).Row + 1 '行の最下に移動

ws2.Range("A" & RowNum).Offset(0, 0) = KeyName
ws2.Range("A" & RowNum).Offset(0, 1) = InDate
ws2.Range("A" & RowNum).Offset(0, 2) = 0
ws2.Range("A" & RowNum).Offset(0, 3) = InStock
ws2.Range("A" & RowNum).Offset(0, 4) = Reason

'入力値初期化
ws1.Range("G12") = 0

End Sub
Sub 品名追加_入力実行()
Attribute 品名追加_入力実行.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 品名追加_入力実行 Macro
'
'変数宣言
Dim KeyName As String '検索ワード
Dim CostPrice As Long
Dim SellPrice As Long
Dim RowNum As Long

Dim SearchRange As Range '検索範囲格納
Dim Result As Range '検索範囲の配列

Set ws1 = Worksheets("入力画面")
Set ws2 = Worksheets("品名追加")
Set ws3 = Worksheets("在庫情報")

'入力値の格納
KeyName = ws2.Range("B6").Value
CostPrice = ws2.Range("E6").Value
SellPrice = ws2.Range("G6").Value

'在庫情報を更新する
If ws3.AutoFilterMode = True Then 'AutoFilterの解除
    ws3.Range("A1").AutoFilter
End If

Set ResultRange = ws3.Range("A:A").Find(KeyName, LookAt:=xlWhole) '最初に一致したRangeを取得

If ResultRange Is Nothing Then '検索結果を判定
    
    RowNum = ws3.Cells(Rows.Count, "A").End(xlUp).Row + 1 '行の最下に移動
    
    ws3.Range("A" & RowNum).Offset(0, 0) = KeyName
    ws3.Range("A" & RowNum).Offset(0, 1) = SellPrice
    ws3.Range("A" & RowNum).Offset(0, 2) = CostPrice
    ws3.Range("A" & RowNum).Offset(0, 3) = 0
    ws3.Range("A" & RowNum).Offset(0, 4) = 0
    
    MsgBox "追加しました。"
    
    Exit Sub

Else
    MsgBox "品名の重複がありました。"
    Exit Sub
End If

'
End Sub
Sub 入庫_入力実行()
'
' 入庫入力実行 Macro
'
'変数宣言
Dim KeyName As String '検索ワード
Dim InDate As Date
Dim InStock As Long
Dim InReson As String
Dim RowNum As Long

Dim Result_Stock As Long
Dim Result_Stock_Sum As Long
Dim Result_Price As Long
Dim Result_Price_Sum As Long

Dim SearchRange As Range '検索範囲格納
Dim Result As Range '検索範囲の配列

Set ws1 = Worksheets("入力画面")
Set ws2 = Worksheets("更新履歴")
Set ws3 = Worksheets("在庫情報")

KeyName = ws1.Range("B8").Value
InDate = ws1.Range("E8").Value
InStock = ws1.Range("G8").Value
Reason = ws1.Range("I8").Value

'在庫情報を更新する

Set ResultRange = ws3.Range("A:A").Find(KeyName, LookAt:=xlWhole) '最初に一致したRangeを取得

If ResultRange Is Nothing Then '検索結果を判定

    MsgBox "検索結果なし"
    
    Exit Sub

Else
    If ws3.AutoFilterMode = True Then 'AutoFilterの解除
        ws3.Range("A1").AutoFilter
    Else
    End If
    
    '在庫データの取得
    Result_Stock = ws3.Range("D" & ResultRange.Row)
    Result_Price = ws3.Range("C" & ResultRange.Row)
    
    '計算
    Result_Stock_Sum = Result_Stock + InStock '結果
    Result_Price_Sum = Result_Stock_Sum * Result_Price '在庫金額
    
    '置換
    ws3.Range("D" & ResultRange.Row) = Result_Stock_Sum
    ws3.Range("E" & ResultRange.Row) = Result_Price_Sum
    
End If

'更新履歴に情報を追加する
If ws2.AutoFilterMode = True Then 'AutoFilterの解除
    ws2.Range("A1").AutoFilter
End If

RowNum = ws2.Cells(Rows.Count, "A").End(xlUp).Row + 1 '行の最下に移動

ws2.Range("A" & RowNum).Offset(0, 0) = KeyName
ws2.Range("A" & RowNum).Offset(0, 1) = InDate
ws2.Range("A" & RowNum).Offset(0, 2) = InStock
ws2.Range("A" & RowNum).Offset(0, 3) = 0
ws2.Range("A" & RowNum).Offset(0, 4) = Reason

'入力値初期化
ws1.Range("G8") = 0

End Sub
Sub 明細の作成()
Attribute 明細の作成.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 明細の作成 Macro
'
'変数宣言
Dim Start_date As Date
Dim Last_date As Date

Dim KeyName As String
Dim Sell_price As Long
Dim Cost_price As Long
Dim Stock_Now As Long

Dim InStock As Long '入庫
Dim SellStock As Long '販売数
Dim ReturnStock As Long '返却
Dim LostStock As Long '紛失
Dim OutStock As Long '欠品
Dim BeforeStock As Long '前在庫
Dim Stock_Status As String '入出庫理由

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

Set ws1 = Worksheets("明細書")
Set ws2 = Worksheets("在庫情報")
Set ws3 = Worksheets("更新履歴")
Set ws4 = Worksheets("印刷ページ")

'明細書印刷ページの初期化
ws4.Cells.ClearContents
ws4.Range("A:I").Borders.LineStyle = xlLineStyleNone

'入力値の格納
Start_date = ws1.Range("B8").Value
Last_date = ws1.Range("E8").Value

'ループ処理
LastRow = ws2.Cells(Rows.Count, 1).End(xlUp).Row 'A列の最終行を取得

For i = 2 To LastRow 'ヘッダーは外す

    'ws2のオートフィルターを解除
    If ws2.AutoFilterMode = True Then 'AutoFilterの解除
        ws2.Range("A1").AutoFilter
    End If

    '在庫情報の取得
    KeyName = ws2.Cells(i, 1)
    Sell_price = ws2.Cells(i, 2)
    Cost_price = ws2.Cells(i, 3)
    Stock_Now = ws2.Cells(i, 4)
    
    '初期値設定
    InStock = 0
    SellStock = 0
    ReturnStock = 0
    LostStock = 0
    OutStock = 0
    BeforeStock = 0
    
    'ws3のオートフィルターを解除
    If ws3.AutoFilterMode = True Then 'AutoFilterの解除
        ws3.Range("A1").AutoFilter
    End If
    
    '更新履歴の処理(オートフィルター)
    ws3.Range("A1").AutoFilter Field:=1, _
    Criteria1:=KeyName
    ws3.Range("A1").AutoFilter Field:=2, _
    Criteria1:=">=" & Start_date, _
    Operator:=xlAnd, _
    Criteria2:="<=" & Last_date
    
    'ヘッダーを除く行数の設定
    n = 2
    
    'フィルター結果の集計
    Do While ws3.Cells(n, 1) <> ""
        
        'Cells(n, 1)のEntireRow.HiddenがFalseなら実行
        If ws3.Cells(n, 1).EntireRow.Hidden = False Then
        
            '在庫状態の取得
            Stock_Status = ws3.Cells(n, 5).Value
            
            '場合分け
            Select Case Stock_Status
            
                Case "入荷"
                
                    InStock = InStock + ws3.Cells(n, 3).Value
                    
                Case "販売数"
                
                    SellStock = SellStock + ws3.Cells(n, 4).Value
                    
                Case "返却"
                
                    ReturnStock = ReturnStock + ws3.Cells(n, 4).Value
                    
                Case "紛失"
                
                    LostStock = LostStock + ws3.Cells(n, 4).Value
                
                Case "欠品"
                
                    OutStock = OutStock + ws3.Cells(n, 4).Value
        
            End Select
        
        End If
        
        n = n + 1
    Loop
    
    '集計結果の作成(前回在庫)
        
    BeforeStock = Stock_Now + SellStock + ReturnStock + LostStock + OutStock - InStock
    
    '明細書の作成
    
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

'ws2のオートフィルターを解除
If ws2.AutoFilterMode = True Then 'AutoFilterの解除
    ws2.Range("A1").AutoFilter
End If

'ws3のオートフィルターを解除
If ws3.AutoFilterMode = True Then 'AutoFilterの解除
    ws3.Range("A1").AutoFilter
End If
    
'明細書枠組みの作成
Set RangeTitle = ws4.Range("A1:I2")
Set RangeDate = ws4.Range("G3:I3")

TodayDate = Date

RangeTitle.MergeCells = True
RangeDate.MergeCells = True

ws4.Range("A1").HorizontalAlignment = xlCenter
ws4.Range("A1").Value = "タイトル"
ws4.Range("A1").Font.Size = 18

ws4.Range("G3").HorizontalAlignment = xlCenter
ws4.Range("G3") = "日付：" & TodayDate

ws4.Range("A4").HorizontalAlignment = xlCenter
ws4.Range("A4").Value = "品名"

ws4.Range("B4").HorizontalAlignment = xlCenter
ws4.Range("B4").Value = "売値"

ws4.Range("C4").HorizontalAlignment = xlCenter
ws4.Range("C4").Value = "仕入値"

ws4.Range("D4").HorizontalAlignment = xlCenter
ws4.Range("D4").Value = "前回在庫"

ws4.Range("E4").HorizontalAlignment = xlCenter
ws4.Range("E4").Value = "新規数"

ws4.Range("F4").HorizontalAlignment = xlCenter
ws4.Range("F4").Value = "販売数"

ws4.Range("G4").HorizontalAlignment = xlCenter
ws4.Range("G4").Value = "小計"

ws4.Range("H4").HorizontalAlignment = xlCenter
ws4.Range("H4").Value = "返却"

ws4.Range("I4").HorizontalAlignment = xlCenter
ws4.Range("I4").Value = "現在庫"

'A列の最終行を取得
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

'小計の合計を追加する。
ws4.Range(ws4.Cells(PrintRow + 1, 7), ws4.Cells(PrintRow + 1, 9)).MergeCells = True
SellSum = WorksheetFunction.Sum(ws4.Range(ws4.Cells(5, 7), ws4.Cells(PrintRow, 7)))
ws4.Cells(PrintRow + 1, 7) = SellSum
ws4.Cells(PrintRow + 1, 7).HorizontalAlignment = xlLeft
ws4.Range(ws4.Cells(PrintRow + 1, 7), ws4.Cells(PrintRow + 1, 9)).Borders(xlEdgeTop).Weight = xlMedium
ws4.Range(ws4.Cells(PrintRow + 1, 7), ws4.Cells(PrintRow + 1, 9)).Borders(xlEdgeLeft).Weight = xlMedium
ws4.Range(ws4.Cells(PrintRow + 1, 7), ws4.Cells(PrintRow + 1, 9)).Borders(xlEdgeBottom).Weight = xlMedium
ws4.Range(ws4.Cells(PrintRow + 1, 7), ws4.Cells(PrintRow + 1, 9)).Borders(xlEdgeRight).Weight = xlMedium

ws4.Cells(PrintRow + 1, 6) = "合計"
ws4.Cells(PrintRow + 1, 6).HorizontalAlignment = xlCenter
ws4.Cells(PrintRow + 1, 6).Borders(xlEdgeTop).Weight = xlMedium
ws4.Cells(PrintRow + 1, 6).Borders(xlEdgeLeft).Weight = xlMedium
ws4.Cells(PrintRow + 1, 6).Borders(xlEdgeBottom).Weight = xlMedium
ws4.Cells(PrintRow + 1, 6).Borders(xlEdgeRight).Weight = xlMedium

ws4.Activate

End Sub

Sub 履歴()
'
' 履歴 Macro
'
    Sheets(5).Select
    Sheets(5).Activate
'
End Sub
Sub チェックボックス作成()

Dim StartX As Single
Dim StartY As Single
Dim EndX As Single
Dim EndY As Single
Dim i As Long
Dim LastRow As Long

    'A列の最終行を取得
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'B列にチェックボックスを作成する
    For i = 1 To LastRow

        With Cells(i, 6)

            'セルの左端
            StartX = .Left

            'セルの上端
            StartY = .Top

            'セルの横幅
            EndX = .Offset(0, 1).Left - .Left

            'セルの高さ
            EndY = .Height

            'チェックボックス作る
            ActiveSheet.CheckBoxes.Add(StartX, StartY, EndX, EndY).Select

            'チェックボックスのテキストを指定
            Selection.Text = ""

            'セルに合わせて移動やサイズを変更する
            Selection.Placement = xlMoveAndSize

        End With

    Next i

End Sub

