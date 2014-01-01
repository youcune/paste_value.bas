Sub paste_value()
    Dim cb As New DataObject
    Dim sel As Range
    Set sel = Selection
     
    If Application.CutCopyMode Then
        sel.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats
    Else
        ' ペーストの起点を決定
        Dim st As Range         ' ペースト起点
        Set st = sel.Range("A1")
         
        ' クリップボードからデータ取得
        Dim c_rows As Variant
        cb.GetFromClipboard
        c_rows = Split(cb.GetText, vbCrLf)
         
        ' 処理中の行/列番号
        Dim i_row As Integer
        i_row = st.Row
        Dim i_col As Integer
        i_col = st.Column
         
        ' ペースト処理
        For i = LBound(c_rows) To UBound(c_rows)
            Dim c_cols As Variant
            c_cols = Split(c_rows(i), vbTab)
            For j = LBound(c_cols) To UBound(c_cols)
                Dim cell As Range
                Set cell = Cells(i_row, i_col)
                With cell
                    .Value = c_cols(j)
                    i_col = i_col + .MergeArea.Columns.Count
                End With
            Next j
             
            ' 改行
            i_col = st.Column
            i_row = i_row + Cells(i_row, i_col).MergeArea.Rows.Count
        Next i
    End If
End Sub
