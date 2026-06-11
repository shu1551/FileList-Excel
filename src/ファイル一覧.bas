Option Base 1
Dim cnt As Long
Dim D1() As Variant, D2() As Variant, D3() As Variant, D4() As Variant, D5() As Variant
Public REKI1 As String
Public REKI2 As String
Public REKI3 As String
Public REKI4 As String
Public REKI5 As String
Public REKI6 As String
Private FSO As Object
Private SKIP_SUB As Boolean



' Windows API宣言 (FindFirstFileW/FindNextFileW) Unicode版
#If VBA7 Then
    Private Declare PtrSafe Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As LongPtr, lpFindFileData As WIN32_FIND_DATA) As LongPtr
    Private Declare PtrSafe Function FindNextFile Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As LongPtr, lpFindFileData As WIN32_FIND_DATA) As Long
    Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As Long
#Else
 ' ↓ #Else(32bit用)が赤いのは正常です。未使用ブランチを示す色で、エラーではありません
    Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long
    Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileW" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
    Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
#End If

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName(0 To 519) As Byte
    cAlternateFileName(0 To 27) As Byte
End Type

Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
#If VBA7 Then
    Private Const INVALID_HANDLE As LongPtr = -1
#Else
    Private Const INVALID_HANDLE As Long = -1
#End If

Sub エクセルを修正リストで連続修正()
    ' --- 入力チェック ---
    If ActiveCell.Column <> 2 Then Exit Sub
    If ActiveCell.Value = "" Then Exit Sub
    If ActiveCell.Offset(0, 4).Value = "" Then Exit Sub

    Dim book1 As Workbook: Set book1 = ThisWorkbook

    ' --- バックアップ先（実行ごとに日時フォルダ）---
    Dim backupDir As String
    backupDir = book1.Path & "\バックアップ_" & Format(Now, "yyyymmdd_hhnnss")

    If vbNo = MsgBox("表示順で連続修正を実行します。" & vbCrLf & _
        "元ファイルは上書き前に下記へバックアップします：" & vbCrLf & _
        backupDir & vbCrLf & vbCrLf & "よろしいですか？", vbYesNo, "処理の確認") Then Exit Sub

    ' バックアップフォルダ作成（作れなければ中止＝保険なしで上書きしない）
    On Error Resume Next
    MkDir backupDir
    On Error GoTo 0
    If Dir(backupDir, vbDirectory) = "" Then
        MsgBox "バックアップフォルダを作成できませんでした。処理を中止します。" & vbCrLf & _
               backupDir, vbCritical
        Exit Sub
    End If
    Dim backupSeq As Long: backupSeq = 0
    Dim bkErr As Long
    Dim targetApp As Object
    Dim book2 As Object
    Dim TargetPath As String, FileName As String
    Dim successCount As Long, failCount As Long, skipCount As Long
    successCount = 0: failCount = 0: skipCount = 0

    ' 修正データを事前に配列に読み込む（高速化）
    Dim lastRowFix As Long
    lastRowFix = book1.Sheets("修正").Cells(Rows.Count, 1).End(xlUp).Row
    If lastRowFix < 2 Then
        MsgBox "修正シートにデータがありません。", vbExclamation
        Exit Sub
    End If
    Dim fixCount As Long
    fixCount = lastRowFix - 1
    Dim arrSheet() As String, arrCell() As String, arrVal() As Variant
    ReDim arrSheet(1 To fixCount)
    ReDim arrCell(1 To fixCount)
    ReDim arrVal(1 To fixCount)
    Dim j As Long
    For j = 1 To fixCount
        arrSheet(j) = book1.Sheets("修正").Range("A" & (j + 1)).Value
        arrCell(j) = book1.Sheets("修正").Range("B" & (j + 1)).Value
        arrVal(j) = book1.Sheets("修正").Range("C" & (j + 1)).Value
    Next j

    ' 隠しExcelを起動
    Set targetApp = CreateObject("Excel.Application")
    targetApp.Visible = False
    targetApp.DisplayAlerts = False

    Application.ScreenUpdating = False

    ' エラー時にも隠しExcelを確実に終了させる
    On Error GoTo CleanupError

    ' アクティブセルから処理を開始
    Do While ActiveCell.Value <> ""
        FileName = ActiveCell.Value
        TargetPath = ActiveCell.Offset(0, 4).Value & "\" & FileName

        If LCase(FileName) Like "*.xls*" Then
            If Dir(TargetPath) <> "" Then

                ' --- 上書き前にバックアップ（失敗したらこのファイルは処理しない）---
                backupSeq = backupSeq + 1
                On Error Resume Next
                Err.Clear
                FileCopy TargetPath, backupDir & "\" & Format(backupSeq, "0000") & "_" & FileName
                bkErr = Err.Number
                On Error GoTo CleanupError
                If bkErr <> 0 Then
                    skipCount = skipCount + 1
                    GoTo NextRow
                End If

                On Error Resume Next
                Set book2 = targetApp.Workbooks.Open(TargetPath)
                On Error GoTo CleanupError

                If Not book2 Is Nothing Then
                    ' ブックを開いた後なら計算モードを設定できる（再計算抑止で高速化）
                    On Error Resume Next
                    targetApp.Calculation = xlCalculationManual
                    On Error GoTo CleanupError
                    ' 配列から修正処理（セル読み取り不要で高速）
                    For j = 1 To fixCount
                        On Error Resume Next
                        book2.Sheets(arrSheet(j)).Range(arrCell(j)).Value = arrVal(j)
                        On Error GoTo CleanupError
                    Next j

                    book2.Close SaveChanges:=True
                    Set book2 = Nothing
                    successCount = successCount + 1
                Else
                    failCount = failCount + 1
                End If
            Else
                skipCount = skipCount + 1
            End If
        Else
            skipCount = skipCount + 1
        End If

NextRow:
        ' 次の可視セルへ移動
        Do
            ActiveCell.Offset(1).Select
        Loop Until Not ActiveCell.EntireRow.Hidden Or ActiveCell.Value = ""
    Loop

    ' 正常終了
    targetApp.Quit
    Set targetApp = Nothing
    Application.ScreenUpdating = True
    MsgBox "処理完了" & vbCrLf & _
           "成功: " & successCount & " 件" & vbCrLf & _
           "失敗: " & failCount & " 件" & vbCrLf & _
           "スキップ: " & skipCount & " 件", vbInformation
    Exit Sub

CleanupError:
    If Not book2 Is Nothing Then
        book2.Close SaveChanges:=False
        Set book2 = Nothing
    End If
    If Not targetApp Is Nothing Then
        targetApp.Quit
        Set targetApp = Nothing
    End If
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "処理済み: " & successCount & " 件" & vbCrLf & _
           Err.Description, vbCritical
End Sub



Sub 選択しているエクセルのマクロ削除()

    ' 変数の宣言
    Dim wBookPathB As String
    Dim wBookB As Workbook
    Dim wCompo As Object ' VBComponent
    Dim wComponents As Object ' VBComponents
    Dim fileRelativePath As String

    ' 画面更新の停止
    Application.ScreenUpdating = False

    ' アクティブセルがExcelファイル名を含まない場合は終了
    If Not LCase(ActiveCell.Value) Like "*.xls*" Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' ファイルパスの構築
    If ActiveCell.Column = 2 Then
        fileRelativePath = ActiveCell.Offset(0, 4).Value & "\" & ActiveCell.Value
    Else
        Application.ScreenUpdating = True
        Exit Sub
    End If

    wBookPathB = fileRelativePath

    ' エラー処理開始
    On Error GoTo ErrorHandler

    ' イベントの無効化
    Application.EnableEvents = False

    ' ブックを開く
    Set wBookB = Application.Workbooks.Open(wBookPathB, ReadOnly:=False)

    With wBookB
        With .VBProject
            Set wComponents = .VBComponents

            ' VBEコンポーネントをループ処理
            For Each wCompo In wComponents
                If wCompo.Type < 4 Then
                    wComponents.Remove wCompo
                ElseIf wCompo.Type = 100 Then
                    With wCompo.CodeModule
                        If .CountOfLines > 0 Then
                            .DeleteLines 1, .CountOfLines
                        End If
                    End With
                End If
            Next wCompo
        End With
    End With

    Call セブクロ

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Unload メニューForm

    Exit Sub

ErrorHandler:
    MsgBox "マクロ削除中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "ファイルパス: " & wBookPathB, vbCritical

    If Not wBookB Is Nothing Then
        wBookB.Close SaveChanges:=False
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
Sub ファイル一覧の実行()
    Dim path01 As String
    Dim idxStart As Long
    Dim idxEnd As Long

    On Error Resume Next
    idxStart = Sheets("全検索").Index
    idxEnd = Sheets("設定").Index
    On Error GoTo 0

    If idxStart = 0 Or idxEnd = 0 Then Exit Sub
    If ActiveSheet.Index <= idxStart Or ActiveSheet.Index >= idxEnd Then Exit Sub

    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Range("A1").Select

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    ActiveSheet.UsedRange.Offset(1).EntireRow.Delete

    If Range("J1").Value <> "" Then
        path01 = Range("J1").Value
    ElseIf Sheets("設定").Range("A2").Value <> "" Then
        path01 = Sheets("設定").Range("A2").Value
    Else
        path01 = ActiveWorkbook.Path
    End If

    cnt = 0
    ReDim D1(1 To 1000)
    ReDim D2(1 To 1000)
    ReDim D3(1 To 1000)
    ReDim D4(1 To 1000)
    ReDim D5(1 To 1000)

    SKIP_SUB = False
    On Error Resume Next
    If Sheets("設定").Range("A16").Value = 1 Then SKIP_SUB = True
    On Error GoTo 0

    Call sagyou(path01)

    If cnt = 0 Then
        MsgBox "データなし"
        Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' 2次元配列に詰めて一括書き込み（Transposeの65,536行制限を回避）
    Dim outArr() As Variant
    Dim r As Long
    ReDim outArr(1 To cnt, 1 To 5)
    For r = 1 To cnt
        outArr(r, 1) = D1(r)
        outArr(r, 2) = D2(r)
        outArr(r, 3) = D3(r)
        outArr(r, 4) = D4(r)
        outArr(r, 5) = D5(r)
    Next r
    Range("B2").Resize(cnt, 5).Value = outArr

    Call sort01
    Call keisen

    Erase D1: Erase D2: Erase D3: Erase D4: Erase D5
    cnt = 0

    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

Sub sagyou(ByVal startPath As String)
    ' === 非再帰・スタック方式 + FindFirstFileW (Unicode対応) ===
    Dim wfd As WIN32_FIND_DATA
    #If VBA7 Then
        Dim hFind As LongPtr
    #Else
        Dim hFind As Long
    #End If
    Dim fName As String
    Dim folderName As String
    Dim fileSize As Double
    Dim searchPath As String

    ' スタック変数
    Dim stackPaths() As String
    Dim stackCount As Long
    Dim stackCapacity As Long
    Dim currentPath As String

    ' スタック初期化
    stackCapacity = 100
    stackCount = 1
    ReDim stackPaths(1 To stackCapacity)
    ' パス末尾に \ を保証
    If Right(startPath, 1) <> "\" Then startPath = startPath & "\"
    stackPaths(1) = startPath

    ' スタックが空になるまでループ（非再帰）
    Do While stackCount > 0
        ' Pop
        currentPath = stackPaths(stackCount)
        stackCount = stackCount - 1

        ' フォルダ名を取得
        Dim pathLen As Long
        pathLen = Len(currentPath)
        If pathLen > 1 Then
            folderName = Mid(currentPath, InStrRev(Left(currentPath, pathLen - 1), "\") + 1)
            If Right(folderName, 1) = "\" Then folderName = Left(folderName, Len(folderName) - 1)
        Else
            folderName = currentPath
        End If

        ' FindFirstFileW で検索開始
        searchPath = currentPath & "*"
        hFind = FindFirstFile(StrPtr(searchPath), wfd)
        If hFind = INVALID_HANDLE Then GoTo NextFolder

        Do
            ' Unicode バイト配列からファイル名を取得
            fName = GetFileName(wfd.cFileName)

            If fName <> "" And fName <> "." And fName <> ".." Then
                If (wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0 Then
                    ' ディレクトリ → スタックに追加（再帰の代わり）
                    If Not SKIP_SUB Then
                            stackCount = stackCount + 1
                            If stackCount > stackCapacity Then
                                stackCapacity = stackCapacity + 100
                                ReDim Preserve stackPaths(1 To stackCapacity)
                            End If
                            stackPaths(stackCount) = currentPath & fName & "\"
                    End If
                Else
                    ' ファイル → 配列に追加
                    If fName <> "desktop.ini" Then
                        cnt = cnt + 1
                        If cnt > UBound(D1) Then
                            Dim newSize As Long
                            newSize = UBound(D1) * 2
                            ReDim Preserve D1(1 To newSize)
                            ReDim Preserve D2(1 To newSize)
                            ReDim Preserve D3(1 To newSize)
                            ReDim Preserve D4(1 To newSize)
                            ReDim Preserve D5(1 To newSize)
                        End If

                        D1(cnt) = fName
                        D2(cnt) = folderName
                        D3(cnt) = FT2Date(wfd.ftLastWriteTime)
                        fileSize = CDec(wfd.nFileSizeHigh) * 4294967296# + CDec(wfd.nFileSizeLow)
                        D4(cnt) = fileSize
                        D5(cnt) = Left(currentPath, Len(currentPath) - 1)

                        If cnt Mod 5000 = 0 Then Application.StatusBar = cnt & " ファイル処理中..."
                    End If
                End If
            End If
        Loop While FindNextFile(hFind, wfd) <> 0

        FindClose hFind

NextFolder:
    Loop
End Sub

Sub sort01()
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Range("B1").Resize(lastRow, 5).Sort _
        Key1:=Range("D2"), Order1:=xlDescending, _
        Header:=xlYes, MatchCase:=False, _
        Orientation:=xlTopToBottom, SortMethod:=xlPinYin
End Sub

Sub keisen()
    Dim lastRow As Long
    ' データ最終行を取得
    lastRow = Cells(Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then Exit Sub ' データがない場合は終了

    Application.ScreenUpdating = False

    ' 連番の作成
    Range("A2").Value = 1
    If lastRow > 2 Then
        Range("A2").AutoFill Destination:=Range("A2:A" & lastRow), Type:=xlFillSeries
    End If

    ' 書式設定
    With Range("A2:F" & lastRow)
        .WrapText = True
    End With

    Range("A2:A" & lastRow).HorizontalAlignment = xlCenter
    Range("D2:D" & lastRow).NumberFormatLocal = "yyyy/mm/dd hh:mm"
    Range("E2:E" & lastRow).NumberFormatLocal = "#,##0_ "
    Range("B2:D" & lastRow).HorizontalAlignment = xlLeft

    Dim maxRow As Long
    Dim maxCol As Long

    maxRow = Cells(Rows.Count, "B").End(xlUp).Row

    maxCol = 6

    If maxRow < 1 Then maxRow = 1

    With Range(Cells(1, 1), Cells(maxRow, maxCol))
        .Borders.LineStyle = True
        .Font.ColorIndex = xlAutomatic
        On Error Resume Next
        .Font.Name = Sheets("設定").Range("F11").Value
        .Font.FontStyle = Sheets("設定").Range("F12").Value
        .Font.Size = Sheets("設定").Range("F13").Value
        On Error GoTo 0
    End With

    Application.ScreenUpdating = True
End Sub

Sub ファイル呼び出し()
    Dim TargetPath As String
    Dim FSO As Object
    Dim Ext As String

    Application.ScreenUpdating = False

    ' FSOオブジェクトの生成
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' 列ごとの処理分岐
    Select Case ActiveCell.Column
        ' ■B列（ファイル名）の場合
        Case 2
            ' パスの結合（F列のパス + \ + B列のファイル名）
            TargetPath = ActiveCell.Offset(0, 4).Value & "\" & ActiveCell.Value

            ' ファイル存在確認
            If FSO.FileExists(TargetPath) Then
                ' 拡張子を取得して小文字に変換
                Ext = LCase(FSO.GetExtensionName(TargetPath))

                ' 拡張子で判定して開き方を変える
                Select Case Ext
                    Case "xls", "xlsx", "xlsm", "xlsb", "csv"
                        Workbooks.Open TargetPath
                    Case Else
                        ' Excel以外は関連付けられたアプリで開く
                        CreateObject("Shell.Application").ShellExecute TargetPath
                End Select
            Else
                MsgBox "ファイルが見つかりません。" & vbCrLf & TargetPath, vbExclamation
            End If

        ' ■C列（参照フォルダ）または F列（直接フォルダ）の場合
        Case 3, 6
            If ActiveCell.Column = 3 Then
                ' C列の場合：F列（3つ右）の値を使用
                TargetPath = ActiveCell.Offset(0, 3).Value
            Else
                ' F列の場合：そのセルの値を使用
                TargetPath = ActiveCell.Value
            End If

            ' フォルダ存在確認
            If TargetPath <> "" Then
                If FSO.FolderExists(TargetPath) Then
                    Shell "C:\Windows\Explorer.exe """ & TargetPath & """", vbNormalFocus
                Else
                    MsgBox "フォルダが見つかりません。" & vbCrLf & TargetPath, vbExclamation
                End If
            End If

    End Select

    Set FSO = Nothing
    Application.ScreenUpdating = True
End Sub
Sub ファイル削除()
    Dim FSO As Object
    Dim ShellApp As Object
    Dim TargetPath As String
    Dim FileName As String
    Dim fullPath As String
    Dim Answer As Integer
    Dim FolderObj As Object
    Dim FileObj As Object

    ' 画面のちらつきを停止
    Application.ScreenUpdating = False

    ' --- 入力チェック ---
    If ActiveCell.Column <> 2 Then GoTo Cleanup
    If ActiveCell.Value = "" Then GoTo Cleanup
    If ActiveCell.Offset(0, 4).Value = "" Then GoTo Cleanup

    ' --- パスの生成 ---
    Set FSO = CreateObject("Scripting.FileSystemObject")

    FileName = ActiveCell.Value
    TargetPath = ActiveCell.Offset(0, 4).Value

    ' パスを正規化して結合
    fullPath = FSO.BuildPath(TargetPath, FileName)

    ' --- 安全対策1：ファイルの存在確認 ---
    If Not FSO.FileExists(fullPath) Then
        MsgBox "指定されたファイルが見つかりません。" & vbCrLf & fullPath, vbExclamation
        GoTo Cleanup
    End If

    ' --- 実行処理（ゴミ箱へ移動） ---
    Set ShellApp = CreateObject("Shell.Application")

    Set FolderObj = ShellApp.Namespace(FSO.GetParentFolderName(fullPath))
    Set FileObj = FolderObj.ParseName(FSO.GetFileName(fullPath))

    If Not FileObj Is Nothing Then
        ShellApp.Namespace(10).MoveHere FileObj
    Else
        MsgBox "ファイルの取得に失敗しました。", vbCritical
        GoTo Cleanup
    End If

    ' --- 安全対策3：削除完了の待機と確認 ---
    DoEvents ' OSに制御を一度戻す

    If FSO.FileExists(fullPath) Then
        MsgBox "ファイルの削除が確認できませんでした。" & vbCrLf & "（キャンセルされたか、ロックされています）", vbInformation
        GoTo Cleanup
    End If

    ' --- セルのクリア ---
    Range(ActiveCell, ActiveCell.Offset(0, 5)).ClearContents

Cleanup:
    Set FSO = Nothing
    Set ShellApp = Nothing
    Set FolderObj = Nothing
    Set FileObj = Nothing
    Application.ScreenUpdating = True
End Sub
Sub 転送()
Dim Copyfilename As String
Dim PathCopy As String
Dim PathName As String
Dim FileName As String
Application.ScreenUpdating = False
c = ActiveCell.Address
Set FSO = CreateObject("Scripting.FileSystemObject")
    If ActiveCell.Row = 1 Then Exit Sub
    If ActiveCell.Column <> 2 Then Exit Sub
    If ActiveCell.Value = "" Then Exit Sub
    FileName = ActiveCell.Value
    PathName = ActiveCell.Offset(0, 4).Value & "\"
    If Sheets("設定").Range("A8").Value <> "" Then
       copysaki = Sheets("設定").Range("A8").Value
       If Right(copysaki, 2) <> "\" Then copysaki = copysaki & "\"
    Else
       copysaki = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\"
    End If
    motoname = PathName & FileName
    Copyfilename = copysaki & FileName
Dim newSaveFilePath As String
newSaveFilePath = create_new_file_path(Copyfilename)
FSO.CopyFile motoname, newSaveFilePath
    Kill motoname
    Set FSO = Nothing
ActiveCell.Offset(0, 0) = Mid(newSaveFilePath, InStrRev(newSaveFilePath, "\") + 1)
ActiveCell.Offset(0, 1) = Mid(Left(copysaki, Len(copysaki) - 1), InStrRev(Left(copysaki, Len(copysaki) - 1), "\") + 1)
ActiveCell.Offset(0, 4) = Left(copysaki, Len(copysaki) - 1)
    With Selection.Font
       .Name = Sheets("設定").Range("F11").Value
        .FontStyle = Sheets("設定").Range("F12").Value
        .Size = Sheets("設定").Range("F13").Value
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
Range(c).Select
Application.ScreenUpdating = True
End Sub
Sub デスクトップにコピー()
Dim Copyfilename As String
Dim PathCopy As String
Dim PathName As String
Dim FileName As String
Application.ScreenUpdating = False
c = ActiveCell.Address
Set FSO = CreateObject("Scripting.FileSystemObject")
    If ActiveCell.Row = 1 Then Exit Sub
    If ActiveCell.Column <> 2 Then Exit Sub
    If ActiveCell.Value = "" Then Exit Sub
     FileName = ActiveCell.Value
    PathName = ActiveCell.Offset(0, 4).Value & "\"
    If Sheets("設定").Range("A5").Value <> "" Then
       copysaki = Sheets("設定").Range("A5").Value
       If Right(copysaki, 2) <> "\" Then copysaki = copysaki & "\"
    Else
       copysaki = CreateObject("WScript.Shell").SpecialFolders.Item("Desktop") & "\"
    End If
    motoname = PathName & FileName
    Copyfilename = copysaki & FileName
Dim newSaveFilePath As String
newSaveFilePath = create_new_file_path(Copyfilename)
FSO.CopyFile motoname, newSaveFilePath
Set FSO = Nothing
Range(c).Select
Application.ScreenUpdating = True
End Sub
Sub ファイル名変換()
    Dim PathName As String
    Dim FileName As String
    Dim fullPath As String
    Dim FSO As Object

    If ActiveCell.Row = 1 Then Exit Sub
    If ActiveCell.Column <> 2 Then Exit Sub
    If ActiveCell.Value = "" Then Exit Sub

    FileName = ActiveCell.Value
    PathName = ActiveCell.Offset(0, 4).Value & "\"

    fullPath = PathName & FileName

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FileExists(fullPath) Then
        MsgBox "変更対象のファイルが見つかりません。" & vbCrLf & fullPath, vbExclamation
        Set FSO = Nothing
        Exit Sub
    End If
    Set FSO = Nothing

    Load ファイル名
    ファイル名.TextBox1.Value = FileName
    ファイル名.Tag = PathName
    ファイル名.Show
End Sub
Sub 検索ワード選択1()
Load 検索フォーム
検索フォーム.StartUpPosition = 0
検索フォーム.Top = Application.Top + ((Application.Height - 検索フォーム.Height) / 2)
検索フォーム.Left = Application.Left + ((Application.Width - 検索フォーム.Width) / 2)
検索フォーム.Show
End Sub
Sub 検索ワード選択2()
Load 検索フォーム3
検索フォーム3.StartUpPosition = 0
検索フォーム3.Top = Application.Top + ((Application.Height - 検索フォーム3.Height) / 2)
検索フォーム3.Left = Application.Left + ((Application.Width - 検索フォーム3.Width) / 2)
検索フォーム3.Show
End Sub
Sub シート名変更()
Load シート名変更フォーム
シート名変更フォーム.StartUpPosition = 0
シート名変更フォーム.Top = Application.Top + ((Application.Height - シート名変更フォーム.Height) / 2)
シート名変更フォーム.Left = Application.Left + ((Application.Width - シート名変更フォーム.Width) / 2)
シート名変更フォーム.Show
End Sub
Sub ALLクリア()
Application.ScreenUpdating = False
c = ActiveCell.Address
If ActiveSheet.FilterMode Then
ActiveSheet.ShowAllData
End If
Application.ScreenUpdating = True
Range("A1").Select
Range(c).Select
End Sub
Sub ユーザーフォーム呼び出し()
Load メニューForm
メニューForm.StartUpPosition = 0
メニューForm.Top = Application.Top + ((Application.Height - メニューForm.Height) / 2)
メニューForm.Left = Application.Left + ((Application.Width - メニューForm.Width) / 2)
メニューForm.Show
End Sub
Sub フォルダ一覧フォーム呼び出し()
Unload メニューForm
フォルダ一覧form.StartUpPosition = 0
フォルダ一覧form.Top = Application.Top + ((Application.Height - フォルダ一覧form.Height) / 2)
フォルダ一覧form.Left = Application.Left + ((Application.Width - フォルダ一覧form.Width) / 2)
フォルダ一覧form.Show
End Sub

Sub お気に入り一覧表示()
    Dim idxStart As Long
    Dim idxEnd As Long

    On Error Resume Next
    idxStart = Sheets("全検索").Index
    idxEnd = Sheets("設定").Index
    On Error GoTo 0

    If idxStart = 0 Or idxEnd = 0 Then Exit Sub

    If ActiveSheet.Index <= idxStart Or ActiveSheet.Index >= idxEnd Then
        Exit Sub
    End If
    Dim wsFav As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Application.ScreenUpdating = False

    Set wsFav = Worksheets("お気に入り")
    lastRow = wsFav.Cells(wsFav.Rows.Count, 1).End(xlUp).Row

    フォルダ一覧form.ListBox1.Clear
    パス.パス.Clear

    REKI6 = "お気に入り"
    REKI1 = "お気に入り"

    If lastRow >= 1 Then
        For i = 1 To lastRow
            フォルダ一覧form.ListBox1.AddItem wsFav.Cells(i, 1).Value
            パス.パス.AddItem wsFav.Cells(i, 2).Value
        Next i
    End If

    With フォルダ一覧form
        .StartUpPosition = 0
        .Top = Application.Top + ((Application.Height - .Height) / 2)
        .Left = Application.Left + ((Application.Width - .Width) / 2)
        .Show
    End With

    Application.ScreenUpdating = True
End Sub

Sub フォルダ一覧表示()
    Dim idxStart As Long
    Dim idxEnd As Long

    On Error Resume Next
    idxStart = Sheets("全検索").Index
    idxEnd = Sheets("設定").Index
    On Error GoTo 0

    If idxStart = 0 Or idxEnd = 0 Then Exit Sub

    If ActiveSheet.Index <= idxStart Or ActiveSheet.Index >= idxEnd Then
        Exit Sub
    End If

    Dim path01 As String
    Dim FSO As Object
    Dim SubFolder As Object

    Application.ScreenUpdating = False

    フォルダ一覧form.ListBox1.Clear
    パス.パス.Clear
    REKI6 = "選択抽出"
    REKI1 = "選択抽出"

    If Range("J1").Value <> "" Then
        path01 = Range("J1").Value
    ElseIf Sheets("設定").Range("A2").Value <> "" Then
        path01 = Sheets("設定").Range("A2").Value
    Else
        path01 = ActiveWorkbook.Path
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    If Not FSO.FolderExists(path01) Then
        MsgBox "指定されたフォルダが見つかりません: " & path01
        Set FSO = Nothing
        Application.ScreenUpdating = True
        Exit Sub
    End If

    For Each SubFolder In FSO.GetFolder(path01).SubFolders
        パス.パス.AddItem SubFolder.Path
        フォルダ一覧form.ListBox1.AddItem SubFolder.Name
    Next SubFolder
    On Error GoTo 0

    Set FSO = Nothing

    With フォルダ一覧form
        .StartUpPosition = 0
        .Top = Application.Top + ((Application.Height - .Height) / 2)
        .Left = Application.Left + ((Application.Width - .Width) / 2)
        .Show
    End With

    Application.ScreenUpdating = True
End Sub

Sub 戻るファルダ一覧表示()
    Application.ScreenUpdating = False

    Call パスログマイナス

    If REKI1 = "お気に入り" Then
        Unload フォルダ一覧form
        Call お気に入り一覧表示
        Exit Sub
    End If

    If REKI1 = "選択抽出" Then
        Unload フォルダ一覧form
        Call フォルダ一覧表示
        Exit Sub
    End If

    Dim path01 As String
    Dim FSO As Object
    Dim TargetFolder As Object
    Dim SubFolder As Object

    path01 = REKI1

    Set FSO = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    Set TargetFolder = FSO.GetFolder(path01)

    If Err.Number <> 0 Then
        MsgBox "フォルダに移動できませんでした。"
        Set FSO = Nothing
        Application.ScreenUpdating = True
        Exit Sub
    End If

    If TargetFolder.SubFolders.Count = 0 Then
        Set FSO = Nothing
        Application.ScreenUpdating = True
        Exit Sub
    End If

    フォルダ一覧form.ListBox1.Clear
    パス.パス.Clear

    For Each SubFolder In TargetFolder.SubFolders
        フォルダ一覧form.ListBox1.AddItem SubFolder.Name
        パス.パス.AddItem SubFolder.Path
    Next SubFolder
    On Error GoTo 0

    Set FSO = Nothing
    Application.ScreenUpdating = True
End Sub
Sub 選択抽出()
    Dim Ix As Long
    Dim path01 As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrorHandler

    If フォルダ一覧form.ListBox1.ListIndex = -1 Then
        Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Ix = フォルダ一覧form.ListBox1.ListIndex
    Unload フォルダ一覧form

    パス.パス.ListIndex = Ix
    path01 = パス.パス.Text

    If Dir(path01, vbDirectory) = "" Then
        MsgBox "フォルダが見つかりません：" & vbCrLf & path01
        GoTo Finalize
    End If

    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    Call シートのクリア
    Range("A1").Select

    cnt = 0
    ReDim D1(1 To 1000)
    ReDim D2(1 To 1000)
    ReDim D3(1 To 1000)
    ReDim D4(1 To 1000)
    ReDim D5(1 To 1000)

    SKIP_SUB = False
    On Error Resume Next
    If Sheets("設定").Range("A16").Value = 1 Then SKIP_SUB = True
    On Error GoTo ErrorHandler

    Call sagyou(path01)

    If cnt = 0 Then
        MsgBox "データなし"
        GoTo Finalize
    End If

    ReDim Preserve D1(1 To cnt)
    ReDim Preserve D2(1 To cnt)
    ReDim Preserve D3(1 To cnt)
    ReDim Preserve D4(1 To cnt)
    ReDim Preserve D5(1 To cnt)

    ' 2次元配列に詰めて一括書き込み（Transposeの65,536行制限を回避）
    Dim outArr() As Variant
    Dim r As Long
    ReDim outArr(1 To cnt, 1 To 5)
    For r = 1 To cnt
        outArr(r, 1) = D1(r)
        outArr(r, 2) = D2(r)
        outArr(r, 3) = D3(r)
        outArr(r, 4) = D4(r)
        outArr(r, 5) = D5(r)
    Next r
    Range("B2").Resize(cnt, 5).Value = outArr

    Call sort01
    Call keisen

Finalize:
    Erase D1: Erase D2: Erase D3: Erase D4: Erase D5
    cnt = 0
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description
    Resume Finalize
End Sub

Sub セブクロ()
    Dim Wb As Workbook
    Dim OtherWbFound As Boolean

    On Error Resume Next
    ActiveWorkbook.Save
    On Error GoTo 0

    OtherWbFound = False
    For Each Wb In Workbooks
        If Wb.Name <> ActiveWorkbook.Name And _
           Not (StrConv(Wb.Name, vbUpperCase) Like "PERSONAL.XLSB" Or _
                StrConv(Wb.Name, vbUpperCase) Like "BOOK.XLSX" Or _
                StrConv(Wb.Name, vbUpperCase) Like "BOOK*.XLSM") Then
            OtherWbFound = True
            Exit For
        End If
    Next Wb

    If OtherWbFound Then
        ActiveWorkbook.Close SaveChanges:=False
    Else
        Application.Quit
    End If

End Sub

Sub 再ファルダ一覧表示()
    Dim Ix As Long
    Dim currentPath As String
    Dim FSO As Object
    Dim TargetFolder As Object
    Dim SubFolder As Object

    Application.ScreenUpdating = False

    If フォルダ一覧form.ListBox1.ListIndex = -1 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    Ix = フォルダ一覧form.ListBox1.ListIndex

    If Ix < パス.パス.ListCount Then
        パス.パス.ListIndex = Ix
        currentPath = パス.パス.Text
    Else
        MsgBox "パスリストの同期エラーです。"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    REKI5 = REKI4
    REKI4 = REKI3
    REKI3 = REKI2
    REKI2 = REKI1
    REKI1 = currentPath

    Set FSO = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    Set TargetFolder = FSO.GetFolder(currentPath)
    If Err.Number <> 0 Then
        MsgBox "フォルダにアクセスできませんでした。" & vbCrLf & "システム制限または削除された可能性があります。", vbExclamation
        Set FSO = Nothing
        Application.ScreenUpdating = True
        Exit Sub
    End If
    On Error GoTo 0

    If TargetFolder.SubFolders.Count = 0 Then
        MsgBox "サブフォルダはありません。"
        Set TargetFolder = Nothing
        Set FSO = Nothing
        Application.ScreenUpdating = True
        Exit Sub
    End If

    フォルダ一覧form.ListBox1.Clear
    パス.パス.Clear

    On Error Resume Next
    For Each SubFolder In TargetFolder.SubFolders
        フォルダ一覧form.ListBox1.AddItem SubFolder.Name
        パス.パス.AddItem SubFolder.Path
    Next SubFolder
    On Error GoTo 0

    Set SubFolder = Nothing
    Set TargetFolder = Nothing
    Set FSO = Nothing

    Application.ScreenUpdating = True
End Sub

Sub パスログクリア()
    REKI5 = ""
    REKI4 = ""
    REKI3 = ""
    REKI2 = ""
    REKI1 = ""
End Sub

Sub パスログプラス()
    REKI5 = REKI4
    REKI4 = REKI3
    REKI3 = REKI2
    REKI2 = REKI1
End Sub

Sub パスログマイナス()
    REKI1 = REKI2
    REKI2 = REKI3
    REKI3 = REKI4
    REKI4 = REKI5
    REKI5 = ""
End Sub
Public Function create_new_file_path(ByVal filePath As String) As String
Dim newFilePath As String
If (Dir(filePath) = "") Then
 newFilePath = filePath
Else
Dim extensionPosition As Long
extensionPosition = InStrRev(filePath, ".")
Dim exceptExtensionFilePaht As String
 Dim extension As String
 If (0 < extensionPosition) Then
 extension = Right(filePath, Len(filePath) - extensionPosition)
 exceptExtensionFilePaht = Left(filePath, extensionPosition - 1)
Else
extension = ""
 exceptExtensionFilePaht = filePath
End If
Dim i As Long
i = 1
Dim num As String
num = i
newFilePath = exceptExtensionFilePaht & "(" & num & ")" & "." & extension
Do While (Dir(newFilePath) <> "")
 i = i + 1
num = i
newFilePath = exceptExtensionFilePaht & "(" & num & ")" & "." & extension
Loop
End If
create_new_file_path = newFilePath
End Function

Sub 全抽出の実行()
    Dim idxStart As Long
    Dim idxEnd As Long
    Dim i As Long
    Dim initialSheet As Worksheet

    Set initialSheet = ActiveSheet

    Application.ScreenUpdating = False

    On Error Resume Next
    idxStart = Sheets("全検索").Index
    idxEnd = Sheets("設定").Index
    On Error GoTo 0

    If idxStart = 0 Or idxEnd = 0 Then
        MsgBox "「全検索」または「設定」シートが見つかりません。", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    If idxStart >= idxEnd - 1 Then
        MsgBox "「全検索」と「設定」の間に処理対象のシートがありません。", vbInformation
        Application.ScreenUpdating = True
        Exit Sub
    End If

    For i = idxStart + 1 To idxEnd - 1
        Sheets(i).Activate

        If ActiveSheet.Visible = xlSheetVisible Then
            Call ファイル一覧の実行
        End If
    Next i

    initialSheet.Activate

    Application.ScreenUpdating = True
End Sub

Private Function GetFileName(ByRef byteArray() As Byte) As String
    ' Unicode バイト配列からファイル名を取得
    Dim tempStr As String
    tempStr = byteArray
    Dim nullPos As Long
    nullPos = InStr(tempStr, vbNullChar)
    If nullPos > 0 Then
        GetFileName = Left$(tempStr, nullPos - 1)
    Else
        GetFileName = tempStr
    End If
End Function

Private Function FT2Date(ft As FILETIME) As Variant
    ' FILETIME → Date 変換（API不要・直接計算方式）
    On Error GoTo ErrHandler

    If ft.dwHighDateTime = 0 And ft.dwLowDateTime = 0 Then
        FT2Date = Null
        Exit Function
    End If

    Dim fileTime64 As Variant
    fileTime64 = CDec(ft.dwHighDateTime) * 4294967296# + CDec(ft.dwLowDateTime)

    ' UTC→ローカル時間補正（日本時間 +9時間 = +324000000000）
    fileTime64 = fileTime64 + CDec(9) * 36000000000#

    ' FILETIMEの基準(1601/1/1)からExcel日付(1899/12/30)への変換
    Dim days As Double
    days = fileTime64 / 864000000000#

    Dim excelDate As Double
    excelDate = days - 109205#

    If excelDate < 1 Or excelDate > 2958465 Then
        FT2Date = Null
    Else
        FT2Date = CDate(excelDate)
    End If
    Exit Function

ErrHandler:
    FT2Date = Null
End Function

Sub WarmUpAPI()
    Dim wfd As WIN32_FIND_DATA
    #If VBA7 Then
        Dim hFind As LongPtr
    #Else
        Dim hFind As Long
    #End If
    Dim searchPath As String

    ' 存在するパスで空振り（1回だけAPIを呼んでDLLをロード）
    searchPath = ThisWorkbook.Path & "\*"
    hFind = FindFirstFile(StrPtr(searchPath), wfd)
    If hFind <> INVALID_HANDLE Then FindClose hFind
End Sub

Sub フォルダ一覧検索()
    フォルダ検索form.StartUpPosition = 0
    フォルダ検索form.Top = Application.Top + ((Application.Height - フォルダ検索form.Height) / 2)
    フォルダ検索form.Left = Application.Left + ((Application.Width - フォルダ検索form.Width) / 2)
    フォルダ検索form.Show
End Sub
