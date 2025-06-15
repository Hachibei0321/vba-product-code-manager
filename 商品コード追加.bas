' ===============================================
' プロシージャ名：商品コード追加
' 作成者：関西のおばちゃん
' 作成日：2025/06/15
' 概要：任意のExcelファイルからワークシートをコピーし、
'       商品コードの照合・追加を行う処理
' ===============================================

Option Explicit

Sub 商品コード追加()
    ' -----------------------------------------------
    ' メイン処理：商品コード追加プロシージャ
    ' -----------------------------------------------
    
    ' 変数の宣言（どんな箱を用意するか決めるで〜）
    Dim 選択ファイルパス As String         ' 選択したファイルのパス
    Dim 選択ワークブック As Workbook        ' 選択したワークブック
    Dim 現在ワークブック As Workbook        ' 今開いてるワークブック
    Dim エラーメッセージ As String         ' エラーが起きた時のメッセージ
    
    ' エラーが起きた時の処理を設定（何かあった時の備えやね）
    On Error GoTo エラー処理
    
    ' 現在開いているワークブックを覚えておく
    Set 現在ワークブック = ThisWorkbook
    
    ' 画面更新を止めて処理を早くする（ちらちらしないようにな）
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' ステップ1：Excelファイルを選択してもらう
    MsgBox "処理を開始するで〜！まずはExcelファイルを選んでや♪", vbInformation, "商品コード追加"
    
    選択ファイルパス = ファイル選択ダイアログ()
    
    ' ファイルが選ばれなかった場合は処理終了
    If 選択ファイルパス = "" Then
        MsgBox "ファイルが選ばれへんかったから、処理をやめるで〜", vbInformation, "処理中止"
        GoTo 処理終了
    End If
    
    ' ステップ2：選択したファイルを開く
    MsgBox "ファイルを開くで〜", vbInformation, "処理中"
    Set 選択ワークブック = Workbooks.Open(選択ファイルパス)
    
    ' ステップ3：ワークシートをコピーする
    MsgBox "ワークシートをコピーするで〜", vbInformation, "処理中"
    Call ワークシートコピー処理(選択ワークブック, 現在ワークブック)
    
    ' ステップ4：商品コードの照合・追加処理
    MsgBox "商品コードの照合を始めるで〜", vbInformation, "処理中"
    Call 商品コード照合処理(選択ワークブック, 現在ワークブック)
    
    ' ステップ5：選択したファイルを閉じる
    選択ワークブック.Close SaveChanges:=False
    
    ' 処理完了のお知らせ
    MsgBox "お疲れさま〜！商品コード追加処理が完了したで♪", vbInformation, "処理完了"
    
処理終了:
    ' 画面更新を元に戻す
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' メモリを解放する（お片付けやね）
    Set 選択ワークブック = Nothing
    Set 現在ワークブック = Nothing
    
    Exit Sub
    
エラー処理:
    ' エラーが起きた時の処理
    エラーメッセージ = "あらあら〜エラーが起きてしもたわ！" & vbCrLf & _
                     "エラー番号：" & Err.Number & vbCrLf & _
                     "エラー内容：" & Err.Description
    
    MsgBox エラーメッセージ, vbCritical, "エラー発生"
    
    ' 開いたファイルがあれば閉じる
    If Not 選択ワークブック Is Nothing Then
        選択ワークブック.Close SaveChanges:=False
    End If
    
    GoTo 処理終了
    
End Sub

' -----------------------------------------------
' ファイル選択ダイアログを表示する関数
' -----------------------------------------------
Function ファイル選択ダイアログ() As String
    
    Dim ファイルダイアログ As FileDialog
    Dim 選択結果 As String
    
    ' ファイル選択ダイアログを作成
    Set ファイルダイアログ = Application.FileDialog(msoFileDialogFilePicker)
    
    ' ダイアログの設定
    With ファイルダイアログ
        .Title = "処理するExcelファイルを選んでや〜"                    ' タイトル
        .Filters.Clear                                                ' フィルタをクリア
        .Filters.Add "Excelマクロファイル", "*.xlsm"                    ' xlsmファイルのみ
        .FilterIndex = 1                                              ' 最初のフィルタを選択
        .AllowMultiSelect = False                                     ' 複数選択は無し
        .InitialFileName = Application.DefaultFilePath                ' 初期フォルダ
        
        ' ダイアログを表示して、OKボタンが押されたら
        If .Show = -1 Then
            選択結果 = .SelectedItems(1)  ' 選択されたファイルのパス
        Else
            選択結果 = ""                 ' キャンセルされた
        End If
    End With
    
    ' 結果を返す
    ファイル選択ダイアログ = 選択結果
    
    ' メモリ解放
    Set ファイルダイアログ = Nothing
    
End Function

' -----------------------------------------------
' ワークシートコピー処理
' -----------------------------------------------
Sub ワークシートコピー処理(コピー元ワークブック As Workbook, コピー先ワークブック As Workbook)
    
    Dim コピー元シート As Worksheet
    Dim 新シート As Worksheet
    Dim シート名配列 As Variant
    Dim i As Integer
    
    ' エラー処理
    On Error GoTo ワークシートコピーエラー
    
    ' コピーするシート名を配列で定義
    シート名配列 = Array("利用率リスト", "管理マスター")
    
    ' 各シートをコピー
    For i = LBound(シート名配列) To UBound(シート名配列)
        
        ' コピー元のシートが存在するかチェック
        If ワークシート存在チェック(コピー元ワークブック, シート名配列(i)) Then
            
            ' シートをコピー
            Set コピー元シート = コピー元ワークブック.Worksheets(シート名配列(i))
            
            ' コピー先に同名シートがあれば削除
            If ワークシート存在チェック(コピー先ワークブック, シート名配列(i)) Then
                Application.DisplayAlerts = False
                コピー先ワークブック.Worksheets(シート名配列(i)).Delete
                Application.DisplayAlerts = True
            End If
            
            ' シートをコピー
            コピー元シート.Copy After:=コピー先ワークブック.Worksheets(コピー先ワークブック.Worksheets.Count)
            
            ' コピーしたシートの名前を設定
            Set 新シート = コピー先ワークブック.Worksheets(コピー先ワークブック.Worksheets.Count)
            新シート.Name = シート名配列(i)
            
            Debug.Print シート名配列(i) & " をコピーしたで〜"
            
        Else
            MsgBox シート名配列(i) & " が見つからへんわ〜", vbExclamation, "シートなし"
        End If
        
    Next i
    
    Exit Sub
    
ワークシートコピーエラー:
    MsgBox "ワークシートのコピーでエラーが起きたで〜" & vbCrLf & Err.Description, vbCritical, "コピーエラー"
    
End Sub

' -----------------------------------------------
' 商品コード照合処理
' -----------------------------------------------
Sub 商品コード照合処理(選択ワークブック As Workbook, 現在ワークブック As Workbook)
    
    Dim 原価リストシート As Worksheet      ' 原価リストのシート
    Dim 料率リストシート As Worksheet      ' 料率リストのシート
    Dim 原価最終行 As Long                 ' 原価リストの最終行
    Dim 料率最終行 As Long                 ' 料率リストの最終行
    Dim i As Long, j As Long               ' ループ用カウンタ
    Dim 照合件数 As Long                   ' 照合できた件数
    
    ' エラー処理
    On Error GoTo 照合処理エラー
    
    ' シートの取得
    Set 原価リストシート = 現在ワークブック.Worksheets("原価リスト")
    Set 料率リストシート = 選択ワークブック.Worksheets("料率リスト")
    
    ' 最終行を取得（B列の最後のデータがある行）
    原価最終行 = 原価リストシート.Cells(原価リストシート.Rows.Count, "B").End(xlUp).Row
    料率最終行 = 料率リストシート.Cells(料率リストシート.Rows.Count, "B").End(xlUp).Row
    
    Debug.Print "原価リスト最終行：" & 原価最終行
    Debug.Print "料率リスト最終行：" & 料率最終行
    
    照合件数 = 0
    
    ' 原価リストの各行をチェック（2行目から開始）
    For i = 2 To 原価最終行
        
        ' 原価リストのB列の値を取得
        Dim 原価Bセル値 As String
        原価Bセル値 = Trim(CStr(原価リストシート.Cells(i, "B").Value))
        
        ' 空白セルなら処理終了
        If 原価Bセル値 = "" Then
            Exit For
        End If
        
        ' 料率リストで同じ値を探す
        For j = 2 To 料率最終行
            
            Dim 料率Bセル値 As String
            料率Bセル値 = Trim(CStr(料率リストシート.Cells(j, "B").Value))
            
            ' 空白セルなら次の原価リスト行へ
            If 料率Bセル値 = "" Then
                Exit For
            End If
            
            ' 値が一致した場合
            If 原価Bセル値 = 料率Bセル値 Then
                
                ' 料率リストのA列の値を原価リストのA列にコピー
                Dim 料率Aセル値 As String
                料率Aセル値 = Trim(CStr(料率リストシート.Cells(j, "A").Value))
                
                原価リストシート.Cells(i, "A").Value = 料率Aセル値
                
                照合件数 = 照合件数 + 1
                
                Debug.Print "照合完了：" & 原価Bセル値 & " → " & 料率Aセル値
                
                ' 一致したら次の原価リスト行へ
                Exit For
                
            End If
            
        Next j
        
    Next i
    
    ' 結果をお知らせ
    MsgBox 照合件数 & " 件の商品コードを追加したで〜♪", vbInformation, "照合結果"
    
    Exit Sub
    
照合処理エラー:
    MsgBox "商品コード照合でエラーが起きたで〜" & vbCrLf & Err.Description, vbCritical, "照合エラー"
    
End Sub

' -----------------------------------------------
' ワークシートの存在チェック関数
' -----------------------------------------------
Function ワークシート存在チェック(対象ワークブック As Workbook, シート名 As String) As Boolean
    
    Dim シート As Worksheet
    
    ' エラーが起きても処理を続ける
    On Error Resume Next
    
    Set シート = 対象ワークブック.Worksheets(シート名)
    
    ' シートが見つかったかどうか
    If Not シート Is Nothing Then
        ワークシート存在チェック = True
    Else
        ワークシート存在チェック = False
    End If
    
    ' エラー処理を元に戻す
    On Error GoTo 0
    
    ' メモリ解放
    Set シート = Nothing
    
End Function
