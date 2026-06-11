Sub 検索ワードの選択()
Load 検索フォーム
検索フォーム.StartUpPosition = 1
検索フォーム.Show
End Sub

Sub 全検索シートで検索実行()
    Dim wsAll As Worksheet
    Dim wsTarget As Worksheet
    Dim CTV As String
    Dim i As Long
    Dim lastRowTarget As Long
    Dim pasteRow As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim rngToCopy As Range

    ' 画面更新と自動計算を停止して高速化
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsAll = Sheets("全検索")

    If ActiveSheet.Name <> "全検索" Then GoTo Finalize

    ' カーソルリセット（SendKeysの代わり）
    Application.GoTo Reference:=wsAll.Range("A1"), Scroll:=True

    ' 既存データのクリア
    Call 全検索シートでシートクリア

    CTV = wsAll.Range("B1").Value
    ' 空白・スペースのみの場合は終了
    If Trim(Replace(CTV, "　", " ")) = "" Then GoTo Finalize

    ' ▼検索ループ開始
    ' 設定シートより左にあるシートを対象とするロジックを維持
    For i = 2 To Sheets("設定").Index - 1
        Set wsTarget = Sheets(i)

        ' データがないシートはスキップ
        lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
        If lastRowTarget <= 1 Then GoTo NextSheet

        ' オートフィルタで検索（A列対象）
        wsTarget.Range("A1").AutoFilter Field:=1, Criteria1:="*" & CTV & "*"

        ' 検索結果（可視セル）があるか確認（ヘッダ以外にあるか）
        If wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row > 1 Then
            ' コピー対象の範囲を取得（ヘッダ除く）
            Set rngToCopy = Nothing
            On Error Resume Next
            Set rngToCopy = wsTarget.Range("A1").CurrentRegion.Offset(1, 0).Resize(wsTarget.Range("A1").CurrentRegion.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
            On Error GoTo 0

            If Not rngToCopy Is Nothing Then
                ' 貼り付け先の行を取得
                pasteRow = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row + 1
                startRow = pasteRow

                ' コピー＆ペースト
                rngToCopy.Copy wsAll.Cells(pasteRow, 1)

                ' 貼り付け後の最終行を取得
                endRow = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row

                ' G列にソースシートのB1セルの値を一括入力（ループさせない）
                If endRow >= startRow Then
                    wsAll.Range(wsAll.Cells(startRow, 7), wsAll.Cells(endRow, 7)).Value = wsTarget.Range("B1").Value
                End If
            End If
        End If

NextSheet:
        ' フィルタ解除
        If wsTarget.FilterMode Then wsTarget.ShowAllData
    Next i

    ' ▼書式設定と連番
    Dim Z As Long
    Z = wsAll.Cells(wsAll.Rows.Count, 1).End(xlUp).Row

    If Z > 1 Then
        ' 連番を一括入力
        wsAll.Range("A2").Value = 1
        If Z > 2 Then
            wsAll.Range("A2").AutoFill Destination:=wsAll.Range("A2:A" & Z), Type:=xlFillSeries
        End If

        ' 書式設定を一括適用（ループさせない）
        With wsAll.Range("A1").CurrentRegion
            .Borders.LineStyle = True
            .Font.ColorIndex = xlAutomatic
            .Font.Name = "Meiryo UI"
            .Font.FontStyle = "標準"
        End With

        ' G列の配置設定
        With wsAll.Range("G2:G" & Z)
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With

        ' ヘッダの色
        wsAll.Range("A1:G1").Font.ThemeColor = xlThemeColorLight1
    End If

Finalize:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ' 完了メッセージが必要ならコメントアウトを外してください
    ' MsgBox "検索完了"
End Sub

Sub 全検索シートでシートクリア()
    If ActiveSheet.Name <> "全検索" Then Exit Sub

    ' カーソルリセット
    Application.GoTo Reference:=ActiveSheet.Range("A1"), Scroll:=True

    Application.ScreenUpdating = False

    ' フィルタ解除
    Call 検索のクリア

    ' データ部分のみ削除（2行目以降）
    Dim lastRow As Long
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    If lastRow > 1 Then
        ActiveSheet.Rows("2:" & lastRow).Delete
    End If

    ' 他シートのフィルタ解除ループ
    Dim i As Long
    ' エラー回避のため設定シートの存在確認を入れるのがベターですが、元のロジックに従います
    On Error Resume Next
    For i = 2 To Sheets("設定").Index - 1
        If Sheets(i).FilterMode Then
            Sheets(i).ShowAllData
        End If
    Next i
    On Error GoTo 0

    Application.ScreenUpdating = True
End Sub

Sub シートのクリア()
' ※このプロシージャが単独で呼ばれる想定の場合
    Application.ScreenUpdating = False
    Call 検索のクリア

    Dim lastRow As Long
    Dim textLastRow As Long
    Dim usedLastRow As Long

    ' 1. A列のデータがある最終行を取得
    textLastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row

    ' 2. 罫線や書式が設定されている最終行を取得（UsedRangeを使用）
    With ActiveSheet.UsedRange
        usedLastRow = .Row + .Rows.Count - 1
    End With

    ' 3. どちらか大きい方を削除対象の最終行とする
    If textLastRow > usedLastRow Then
        lastRow = textLastRow
    Else
        lastRow = usedLastRow
    End If

    ' 4. 削除と罫線クリアの実行
    If lastRow > 1 Then
        ' 先に行を削除（これで基本的には消えます）
        ActiveSheet.Rows("2:" & lastRow).Delete

        ' ★追加部分：削除後に繰り上がってきた行や、残存する範囲の罫線を念のため消去
        ' （データ量が多いと重くなるため、目に見える範囲＋α程度に限定しても良いですが、ここでは確実性を優先）
        On Error Resume Next
        ActiveSheet.Rows("2:" & lastRow).Borders.LineStyle = xlNone
        On Error GoTo 0
    End If

    Application.ScreenUpdating = True
End Sub

Sub 検索のクリア()
    ' シンプルにフィルタ解除のみを行う
    If ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
End Sub
Sub マクロの一覧()
Application.ScreenUpdating = False
マクロフォーム.ListBox1.Clear
    For i = 1 To ActiveWorkbook.VBProject.VBComponents.Count
     If ActiveWorkbook.VBProject.VBComponents(i).Type <> 1 Then GoTo KOKO
      With ActiveWorkbook.VBProject.VBComponents(i).CodeModule
        proc = ""
        For j = 1 To .CountOfLines
          If proc <> .ProcOfLine(j, 0) Then
            proc = .ProcOfLine(j, 0)
            マクロフォーム.ListBox1.AddItem proc
          End If
        Next j
      End With
KOKO:
    Next i
On Error Resume Next
Load マクロフォーム
マクロフォーム.StartUpPosition = 0
マクロフォーム.Top = Application.Top + ((Application.Height - マクロフォーム.Height) / 2)
マクロフォーム.Left = Application.Left + ((Application.Width - マクロフォーム.Width) / 2)
マクロフォーム.ListBox1.Selected(0) = True
マクロフォーム.Show vbModeless
Application.ScreenUpdating = True
End Sub
