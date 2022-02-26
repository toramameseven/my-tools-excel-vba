
Option Explicit

Sub checkExistOrNot()
    
    'select range
    'output which column has which item or not.

    '' original data
    Dim dataColumns As Long
    Dim dataRows As Long
    Dim rngData As Range
    
    
    Set rngData = Selection
    dataColumns = rngData.columns.Count
    dataRows = rngData.Rows.Count
    
    If dataColumns < 2 Then
        MsgBox "Select 2 or more columns"
        Exit Sub
    End If
    
    Dim data As Variant
    data = rngData
    
    'get title
    Dim title As Variant
    ReDim title(0 To 1, 0 To dataColumns)
    If rngData.Resize(1).row > 1 Then
        title = rngData.Resize(1).Offset(-1)
    End If
    
    
    '' output range
    Dim rangeDiff As Range
    Set rangeDiff = Application.InputBox(prompt:="", title:="Select the topleft cell for Output.", Type:=8)
    If WorksheetFunction.CountA(rangeDiff.Resize(dataRows * dataColumns + 1, dataColumns + 1)) <> 0 Then
        MsgBox "Output range is not Empty."
        Exit Sub
    End If

    
    '' dictionary
    Dim dic() As Object
    ReDim dic(1 To dataColumns)
    Dim c As Long
    
    For c = 1 To dataColumns
        Set dic(c) = CreateObject("Scripting.Dictionary")
    Next
    
    Dim dicDiff As Object
    Set dicDiff = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim v As String

    For c = 1 To dataColumns
        For i = 1 To dataRows
            v = data(i, c)
            If v <> "" Then
                If Not dic(c).Exists(v) Then
                    dic(c).Add v, v
                Else
                    MsgBox v & "is duplicate. column:" & c
                    Exit Sub
                End If
                
                If Not dicDiff.Exists(v) Then
                    dicDiff.Add v, v
                End If
            End If
        Next
    Next
    
    
    '' sort key
    Dim allKeys As Variant  '
    allKeys = dicDiff.Keys
    
    Dim diffArrayList As Object
    Set diffArrayList = CreateObject("System.Collections.ArrayList")
    
    For i = 0 To UBound(allKeys)
        diffArrayList.Add (allKeys(i))
    Next
    diffArrayList.Sort
    
    
    '' make result and output the result
    Dim outValue As Variant
    outValue = rangeDiff.Resize(UBound(allKeys) + 2, dataColumns + 1)
    
    For i = 0 To UBound(allKeys)
        outValue(i + 2, 1) = diffArrayList(i)
        For c = 1 To dataColumns
            outValue(i + 2, c + 1) = IIf(dic(c).Exists(diffArrayList(i)), 1, 0)
        Next
    Next
    
    rangeDiff.Resize(UBound(allKeys) + 2, dataColumns + 1) = outValue
    If rngData.Resize(1).row > 1 Then
        rangeDiff.Resize(1, dataColumns).Offset(0, 1) = title
    End If
End Sub
Public Sub furiganaDel()
    Selection.Characters.PhoneticCharacters = ""
End Sub


Private Sub saveutf8()
    'Dim writeStream As ADODB.Stream
    'Microsoft ActiveX Data Objects 2.5 Libraryと
    Dim writeStream As Object
    ' 文字コードを指定してファイルをオープン
    Set writeStream = CreateObject("ADODB.Stream")
    writeStream.Charset = "UTF-8"
    writeStream.Open

    ' バッファに出力
    writeStream.WriteText Cells(1, 1), 1 'adWriteLine '1
    writeStream.WriteText Cells(2, 1), 1  'adWriteLine '1
    writeStream.WriteText Cells(3, 1), 1  'adWriteLine '1
    ' ファイルに書き込み
    writeStream.SaveToFile "C:\test.txt", 2 'adSaveCreateOverWrite:2

    ' ファイルをクローズ
    writeStream.Close
    Set writeStream = Nothing
End Sub


Private Sub saveutf8withoutBom()
    Exit Sub
    Dim tmpWriteStream As Object  'As ADODB.Stream
    Dim readStream As Object  'ADODB.Stream
    Dim writeStream As Object  'ADODB.Stream
    Dim tmpFile As String

    ' 一時ファイル
    tmpFile = "C:\tmp_file.txt"

    ' 文字コードを指定してファイルをオープン（一時ファイル）
    Set tmpWriteStream = CreateObject("ADODB.Stream")
    tmpWriteStream.Charset = "UTF-8"
    tmpWriteStream.Open

    ' バッファに出力
    tmpWriteStream.WriteText "1行目", 1 'adWriteLine '1
    tmpWriteStream.WriteText "2行目", 1 'adWriteLine '1
    ' 一時ファイルに書き込み
    tmpWriteStream.SaveToFile tmpFile, 2 'adSaveCreateOverWrite:2

    ' 一時ファイルをクローズ
    tmpWriteStream.Close
    Set tmpWriteStream = Nothing

    ' ここからBOM除去処理
    ' 一時ファイルの4バイト目から読み込む
    Set readStream = CreateObject("ADODB.Stream")
    readStream.Open
    readStream.Type = 1 'adTypeBinary:1
    readStream.LoadFromFile (tmpFile)
    readStream.Position = 3

    ' 出力ファイルをオープン
    Set writeStream = CreateObject("ADODB.Stream")
    writeStream.Open
    writeStream.Type = 1 'adTypeBinary:1
    ' 読み込んだデータをそのままファイルに出力する（4バイト目以降を出力）
    writeStream.Write (readStream.Read(-1)) '-1 adReadAll
    writeStream.SaveToFile "C:\test.txt", 2 'adSaveCreateOverWrite:2

    ' ファイルをクローズ
    writeStream.Close
    Set writeStream = Nothing
    readStream.Close
    Set readStream = Nothing

    ' 一時ファイル削除
    Kill tmpFile
End Sub


''Sub maketestpic()
''    Dim i As Long
''    Dim j As Long
''
''    Dim rng As Range
''    Set rng = Range("A1")
''
''    For i = 0 To 9
''        For j = 0 To 4
''            rng.Offset(i, j).Select
''            Call RenderPictureAndSave
''        Next j
''    Next i
''End Sub


'---------------------------------------------------------------------------------------
' Procedure : duplicateCheck
' Author    : toramame
' Date      : 2016/12/13
' Purpose   : 重複チェック機能  複数列のデータを、列ごとに重複を取り去る
'---------------------------------------------------------------------------------------
'
Sub duplicateCheck()

    On Error GoTo duplicateCheck_Error
    
    Dim myDic As Object, myKey As Variant, myCellPos As Variant
    Dim c As Variant, varData As Variant
    Dim rngCommnet As Range
    
    Dim myrows As Long
    Dim mycolumns As Long
    
    Dim isOutputCellPos As Boolean
    isOutputCellPos = True
    
    Dim myRange As Range
    Set myRange = Range(Selection.Address)
    
    myrows = myRange.Rows.Count
    mycolumns = myRange.columns.Count
    
    Dim distRange As Range
    Set distRange = Application.InputBox(prompt:="", title:="コピー先", Type:=8)
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.StatusBar = "重複チェック"
    
    
    
    Dim i As Long
    Dim j As Long
    Dim strcellpos As String
    For i = 1 To mycolumns
        Set myDic = CreateObject("Scripting.Dictionary")
        varData = myRange.Resize(myrows, 1).Offset(0, i - 1).Value
        
        For j = 1 To myrows
            c = varData(j, 1)
            strcellpos = "(" & j & "," & i & ") "
            If Not c = Empty Then
                If Not myDic.Exists(c) Then
                    myDic.Add c, strcellpos  'cが Keyで、項目が にセル位置を追加していく
                Else
                    myDic.Item(c) = myDic.Item(c) & strcellpos
                End If
            End If
        Next j
        
        '配列が返ってくる
        myKey = myDic.Keys
        If myDic.Count > 0 Then
            If WorksheetFunction.CountA(distRange.Resize(myDic.Count, 1).Offset(0, i - 1)) <> 0 Then
                MsgBox "出力左記が空欄でありません"
                Exit For
            Else
                distRange.Resize(myDic.Count, 1).Offset(0, i - 1) = Application.WorksheetFunction.Transpose(myKey)
                
                If isOutputCellPos Then
                    myCellPos = myDic.Items
                    For j = 1 To UBound(myCellPos) + 1
                        Set rngCommnet = distRange.Offset(j - 1, i - 1).Resize(1, 1)
                        If TypeName(rngCommnet.Comment) = "Nothing" Then
                            rngCommnet.AddComment
                            rngCommnet.Comment.Visible = False
                            rngCommnet.Comment.text text:=myCellPos(j - 1)
                        Else
                            rngCommnet.Comment.text text:=myCellPos(j - 1)
                        End If
                    Next j
                End If
            End If
        End If
    Next i
    
    Set myDic = Nothing

    On Error GoTo 0
    GoTo duplicateCheck_Normal_exit

duplicateCheck_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure duplicateCheck of Module modIROIRO"

duplicateCheck_Normal_exit:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Sub diffTwoLine()
    Dim range1 As Range
    Dim range2 As Range
    Dim rangeDiff As Range
    
    Dim rows1 As Long
    Dim rows2 As Long
    
    Set range1 = Selection
    If range1.columns.Count > 1 Then
        MsgBox "2列以上は選択できません"
        Exit Sub
    End If
    rows1 = range1.Rows.Count
    
    Set range2 = Application.InputBox(prompt:="", title:="2列目を指定する", Type:=8)
    If range2.columns.Count > 1 Then
        MsgBox "2列以上は選択できません"
        Exit Sub
    End If
    rows2 = range2.Rows.Count
    
    Set rangeDiff = Application.InputBox(prompt:="", title:="出力先を指定する", Type:=8)
    
    If WorksheetFunction.CountA(rangeDiff.Resize(rows1 + rows2, 3)) <> 0 Then
        MsgBox "比較結果出力先が空欄でありません"
        Exit Sub
    End If


    Dim r(1 To 2) As Variant
    r(1) = range1
    r(2) = range2
    
    Dim dic1 As Object
    Dim dic2 As Object
    Dim dicDiff As Object
    Set dic1 = CreateObject("Scripting.Dictionary")
    Set dic2 = CreateObject("Scripting.Dictionary")
    Set dicDiff = CreateObject("Scripting.Dictionary")
    
    
    Dim i As Long
    Dim v As String


    
    For i = 1 To rows1
        v = r(1)(i, 1)
        If Not dic1.Exists(v) Then
            dic1.Add v, v
            dicDiff.Add v, v
        Else
            MsgBox v & "が重複しています。列1"
            Exit Sub
        End If
    Next
        
    For i = 1 To rows2
        v = r(2)(i, 1)
        If Not dic2.Exists(v) Then
            dic2.Add v, v
            If Not dicDiff.Exists(v) Then
                dicDiff.Add v, v
            End If
        Else
            MsgBox v & "が重複しています。列2"
            Exit Sub
        End If
    Next
    
    Dim allKeys As Variant  '
    allKeys = dicDiff.Keys
    
    Dim diffArrayList As Object
    Set diffArrayList = CreateObject("System.Collections.ArrayList")
    
    For i = 0 To UBound(allKeys)
        diffArrayList.Add (allKeys(i))
    Next
    
    diffArrayList.Sort
    
    Dim outValue As Variant
    outValue = rangeDiff.Resize(UBound(allKeys) + 1, 3)
    
    For i = 0 To UBound(allKeys)
        outValue(i + 1, 1) = diffArrayList(i)
        outValue(i + 1, 2) = dic1.Exists(diffArrayList(i))
        outValue(i + 1, 3) = dic2.Exists(diffArrayList(i))
    Next
    rangeDiff.Resize(UBound(allKeys) + 1, 3) = outValue
End Sub







'---------------------------------------------------------------------------------------
' Procedure : duplicateCheckMatrix
' Author    : toramame
' Date      : 2016/12/14
' Purpose   : 行列データから、重複を避けて1列にする
'---------------------------------------------------------------------------------------
'
Sub duplicateCheckMatrix(ByVal strAddressToOutput As String)
    Dim myDic As Object, myKey As Variant
    Dim c As Variant, varData As Variant
    Dim myrows As Long
    Dim mycolumns As Long
    Dim myRange As Range
    
    Dim singlewords As Long
    Dim duplicatewords As Long
    
    
    On Error GoTo duplicateCheckMatrix_Error
    Set myRange = Range(Selection.Address)
    
    myrows = myRange.Rows.Count
    mycolumns = myRange.columns.Count
    
    Dim distRange As Range
    If strAddressToOutput = "" Then
        Set distRange = Application.InputBox(prompt:="", title:="コピー先", Type:=8)
    Else
        Set distRange = Range(strAddressToOutput)
    End If
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.StatusBar = ""
    
    Dim i As Long
    Dim j As Long
    
    Set myDic = CreateObject("Scripting.Dictionary")
    
    varData = myRange.Resize.Value
    Dim strcellpos As String
    Dim ii As Long
    For i = 1 To myrows
        For j = 1 To mycolumns
        c = CStr(varData(i, j))
        
        
        ii = i 'IIf(485 < i, i - 485, i)
        
        
        'strcellpos = "" & i & "," & j & " "
        strcellpos = ii & ", "
        
        If Not c = Empty Then
            If Not myDic.Exists(c) Then
                myDic.Add c, strcellpos  'cが Keyで、項目が にセル位置を追加していく
                singlewords = singlewords + 1
            Else
                myDic.Item(c) = myDic.Item(c) & strcellpos
                duplicatewords = duplicatewords + 1
            End If
            
        End If
        Next j
    Next i
    
    '配列が返ってくる
    myKey = myDic.Keys
    
    Dim myItems As Variant
    
    If WorksheetFunction.CountA(distRange.Resize(myDic.Count, 2)) <> 0 Then
        MsgBox "出力左記が空欄でありません"
    Else
        distRange.Resize(myDic.Count, 1) = Application.WorksheetFunction.Transpose(myKey)
        
       
        
        Dim myCellPos As Variant
        Dim rngCommnet As Range
        If True Then
            myCellPos = myDic.Items
            For j = 1 To UBound(myCellPos) + 1
            
                Set rngCommnet = distRange.Offset(j - 1, 0).Resize(1, 1)
                
'' コメントで位置設定
'                If TypeName(rngCommnet.Comment) = "Nothing" Then
'                    rngCommnet.AddComment
'                    rngCommnet.Comment.Visible = False
'                    rngCommnet.Comment.Text Text:=myCellPos(j - 1)
'                Else
'                    rngCommnet.Comment.Text Text:=myCellPos(j - 1)
'                End If
               
''別セルに位置設定
                rngCommnet.Offset(0, 1) = myCellPos(j - 1)

            Next j
        End If
        
    End If
    Application.GoTo distRange, True
    MsgBox "単独:" & singlewords & " & 重複:" & duplicatewords
    
    On Error GoTo 0
    GoTo duplicateCheckMatrix_Normal_exit

duplicateCheckMatrix_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure duplicateCheckMatrix of Module modIROIRO"

duplicateCheckMatrix_Normal_exit:
    Set myDic = Nothing
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub




'---------------------------------------------------------------------------------------
' Procedure : KeyValueSet
' Author    : toramame
' Date      : 2016/12/14
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub KeyValueSet()
    Dim myDic As Object, myKey As Variant
    Dim c As Variant
    
    Dim myrows As Long
    Dim mycolumns As Long
    
    Dim myRange As Range
    
    'On Error GoTo KeyValueSet_Error

    Set myRange = Range(Selection.Address)
    
    myrows = myRange.Rows.Count
    mycolumns = myRange.columns.Count
    
    
    If mycolumns <> 1 Then
        MsgBox "選択列を1列にしてください。"
        GoTo KeyValueSet_Normal_exit
    End If
    
    Dim rngValue As Range
    Set rngValue = Application.InputBox(prompt:="", title:="キーに対する、値の先頭行を選択してください", Type:=8)
    
    
    Dim distRange As Range
    Set distRange = Application.InputBox(prompt:="", title:="適用先のキー範囲一列で", Type:=8)
    
    Dim distRangeOutput As Range
    Set distRangeOutput = Application.InputBox(prompt:="", title:="値の出力先を 1セル選択", Type:=8)
    
    Dim i As Long
    Dim varData As Variant
    Dim varValue As Variant
    Dim itemData As Variant
    Set myDic = CreateObject("Scripting.Dictionary")
    varData = myRange.Value
    varValue = rngValue.Resize(myrows, 1)
   
    For i = 1 To myrows
        c = varData(i, 1)
        If Not c = Empty Then
            c = CStr(c)
            If Not myDic.Exists(c) Then
                myDic.Add c, varValue(i, 1) 'varData(i, 2)
            Else
                MsgBox "キーが重複しています::" & c
                GoTo KeyValueSet_Normal_exit
            End If
        End If
    Next
    
    '配列が返ってくる
    Dim j As Long
    Dim itemval As String
    Dim keytemp As String
    'If WorksheetFunction.CountA(distRange.Offset(0, 1)) <> 0 Then
    If WorksheetFunction.CountA(distRangeOutput.Resize(distRange.row, 1)) <> 0 Then
        MsgBox "出力先が空欄でありません" & distRange.Offset(0, 1).Address
        GoTo KeyValueSet_Normal_exit
    Else
        For j = 1 To distRange.Rows.Count
            keytemp = CStr(distRange.Cells(j, 1))
            If myDic.Exists(keytemp) Then
                itemval = myDic.Item(keytemp)
            Else
                itemval = "###### No Item #####"
            End If
            'distRange.Offset(0, 1).Cells(j, 1) = itemval
            distRangeOutput.Offset(0, 0).Cells(j, 1) = itemval
        Next j
    End If

    On Error GoTo 0
    GoTo KeyValueSet_Normal_exit

KeyValueSet_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure KeyValueSet of Module modIROIRO"

KeyValueSet_Normal_exit:
    Set myDic = Nothing
    
End Sub

Sub KeyValueSet2()
    ' 複数列バリューに対応させた
     Dim myDic As Object, myKey As Variant
     Dim c As Variant

     Dim myrows As Long
     Dim mycolumns As Long

     Dim myRange As Range

     On Error GoTo KeyValueSet_Error

     Set myRange = Range(Selection, Selection.End(xlDown))

     myrows = myRange.Rows.Count
     mycolumns = myRange.columns.Count

     If mycolumns <> 1 Then
         MsgBox "選択列を１列にしてください。"
         GoTo KeyValueSet_Normal_exit
     End If

     Dim valueRange As Range
     Set valueRange = Application.InputBox(prompt:="", title:="キーに対する値の列の最初の行を選択", Type:=8)
     Dim valuecols As Long
     valuecols = valueRange.columns.Count

     Dim distKeyRange As Range
     Set distKeyRange = Application.InputBox(prompt:="", title:="適用先の値の列の初め", Type:=8)
     Set distKeyRange = Range(distKeyRange, distKeyRange.End(xlDown))

     Dim distValueRange As Range
     Set distValueRange = Application.InputBox(prompt:="", title:="適用先のバリューの最初の行", Type:=8)

     Dim i As Long
     Dim j As Long
     Dim varData As Variant
     Dim varValue As Variant
     Dim itemData As Variant
     Set myDic = CreateObject("Scripting.Dictionary")
     varData = myRange.Value
     varValue = valueRange.Resize(myrows, valuecols).Value

     Dim vararray() As Variant
     For i = 1 To myrows
         c = varData(i, 1)
         If Not c = Empty Then
             If Not myDic.Exists(c) Then
                 ReDim vararray(1 To valuecols)
                 For j = 1 To valuecols
                     vararray(j) = varValue(i, j)
                 Next j
                 myDic.Add CStr(c), vararray
             Else
                 MsgBox "キーが重複しています::" & c
                 GoTo KeyValueSet_Normal_exit
             End If
         End If
     Next


     Dim itemval As String
     Dim keytemp As String
     If WorksheetFunction.CountA(distValueRange.Resize(distKeyRange.Rows.Count, valuecols)) <> 0 Then
         MsgBox "出力先が空欄でありません"
         GoTo KeyValueSet_Normal_exit
     Else
         For i = 1 To distKeyRange.Rows.Count
             keytemp = CStr(distKeyRange.Cells(i, 1))
             If myDic.Exists(keytemp) Then
                 For j = 1 To valuecols
                     distValueRange.Resize(myrows, valuecols).Cells(i, j) = myDic.Item(keytemp)(j)
                 Next j
             Else
                 For j = 1 To valuecols
                     distValueRange.Resize(myrows, valuecols).Cells(i, j) = "#no val#"
                 Next j
             End If

         Next i
     End If

     On Error GoTo 0
     GoTo KeyValueSet_Normal_exit

KeyValueSet_Error:
     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure KeyValueSet2 of Module modIROIRO"

KeyValueSet_Normal_exit:
     Set myDic = Nothing

End Sub


Sub KeyValueSet33()
    ' キーとキーが それぞれあるかどうか確認
     Dim myDic As Object, myKey As Variant
     Dim c As Variant

     Dim myrows As Long
     Dim mycolumns As Long

     Dim myRange As Range

     On Error GoTo KeyValueSet_Error

     Set myRange = Range(Selection, Selection.End(xlDown))

     myrows = myRange.Rows.Count
     mycolumns = myRange.columns.Count

     If mycolumns <> 1 Then
         MsgBox "選択列を１列にしてください。"
         GoTo KeyValueSet_Normal_exit
     End If

     Dim myRange2 As Range
     Set myRange2 = Application.InputBox(prompt:="", title:="比較対象の列", Type:=8)
     Set myRange2 = Range(myRange2, myRange2.End(xlDown))

     Dim distRange As Range
     Set distRange = Application.InputBox(prompt:="", title:="適用先の1", Type:=8)
     Set distRange = Range(distRange, distRange.End(xlDown))

     Dim distRange2 As Range
     Set distRange2 = Application.InputBox(prompt:="", title:="適用先の2", Type:=8)
     Set distRange2 = Range(distRange2, distRange2.End(xlDown))
     

     Dim i As Long
     Dim j As Long
     Dim varData As Variant
     Set myDic = CreateObject("Scripting.Dictionary")
     varData = myRange.Value

     Dim vararray() As Variant
     For i = 1 To myrows
         c = varData(i, 1)
         If Not c = Empty Then
             If Not myDic.Exists(c) Then
                 myDic.Add CStr(c), "Data1"
             Else
                 MsgBox "キーが重複しています::" & c
                 GoTo KeyValueSet_Normal_exit
             End If
         End If
     Next

     Dim itemval As String
     Dim keytemp As String
     Dim tvalaa As String
     If WorksheetFunction.CountA(distRange2.Resize(myRange2.Rows.Count, 1)) <> 0 Then
         MsgBox "出力先が空欄でありません"
         GoTo KeyValueSet_Normal_exit
     Else
         For i = 1 To myRange2.Rows.Count
             keytemp = CStr(myRange2.Cells(i, 1))

             
             If myDic.Exists(keytemp) Then
                tvalaa = "N1"
             Else
                 tvalaa = "#Noval#"
             End If
             distRange2.Resize(myrows, 1).Cells(i, 1) = tvalaa
         Next i
     End If
     
     
     '''''
     myrows = myRange2.Rows.Count
     Set myDic = CreateObject("Scripting.Dictionary")
     varData = myRange2.Value
     
      For i = 1 To myrows
         c = varData(i, 1)
         If Not c = Empty Then
             If Not myDic.Exists(c) Then
                 myDic.Add CStr(c), "Data2"
             Else
                 MsgBox "キーが重複しています::" & c
                 GoTo KeyValueSet_Normal_exit
             End If
         End If
     Next

     If WorksheetFunction.CountA(distRange.Resize(myRange.Rows.Count, 1)) <> 0 Then
         MsgBox "出力先が空欄でありません"
         GoTo KeyValueSet_Normal_exit
     Else
         For i = 1 To myRange.Rows.Count
             keytemp = CStr(myRange.Cells(i, 1))
             If myDic.Exists(keytemp) Then
                tvalaa = "N2"
             Else
                 tvalaa = "#Noval#"
             End If
             distRange.Resize(myrows, 1).Cells(i, 1) = tvalaa
         Next i
     End If
     
     
    
     On Error GoTo 0
     GoTo KeyValueSet_Normal_exit

KeyValueSet_Error:
     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure KeyValueSet2 of Module modIROIRO"

KeyValueSet_Normal_exit:
     Set myDic = Nothing

End Sub
Sub test()


    Dim r1 As Range
    Dim r2 As Range
    Dim r3 As Range
    Dim r4 As Range
    
    Set r1 = Worksheets("201811不具合処理").Range("C2")
    Set r2 = Worksheets("201811不具合処理").Range("A2")
    
    Set r3 = Worksheets("不具合チケット80000_追加情報").Range("A2")
    
    
    Set r4 = Worksheets("不具合チケット80000_追加情報").Range("K2")
    Range(r4, r4.End(xlDown)).Clear
    
    KeyValueSet3 r1, r2, r3, r4
    
    
    
    Set r1 = Worksheets("201811不具合処理").Range("C2")
    Set r2 = Worksheets("201811不具合処理").Range("A2")
    
    Set r3 = Worksheets("i-QLinksチケットALL").Range("A2")
    Set r4 = Worksheets("i-QLinksチケットALL").Range("K2")
    Range(r4, r4.End(xlDown)).Clear

    KeyValueSet3 r1, r2, r3, r4
    
    
End Sub


Sub test2()


    Dim r1 As Range
    Dim r2 As Range
    Dim r3 As Range
    Dim r4 As Range
    
    Set r1 = Worksheets("不具合チケット80000_追加情報 (2)").Range("A2")
    Set r2 = Worksheets("不具合チケット80000_追加情報 (2)").Range("L2")
    
    Set r3 = Worksheets("201811不具合処理").Range("C2")
    
    'Reademeのバージョン
    Set r4 = Worksheets("201811不具合処理").Range("N2")
    r4.Resize(Rows.Count - 1).Clear
    
    KeyValueSet3 r1, r2, r3, r4
    
End Sub

Sub test3()


    Dim r1 As Range
    Dim r2 As Range
    Dim r3 As Range
    Dim r4 As Range
    
    Set r1 = Worksheets("不具合チケット80000_追加情報 (2)").Range("A2")
    Set r2 = Worksheets("不具合チケット80000_追加情報 (2)").Range("E2:G2")
    
    Set r3 = Worksheets("201811不具合処理").Range("C2")
    
    'どこの不具合か
    Set r4 = Worksheets("201811不具合処理").Range("Q2").Resize(1, 3)
    r4.Resize(Rows.Count - 1).Clear
    

    KeyValueSet3 r1, r2, r3, r4
    
End Sub



Sub KeyValueSet3(r1 As Range, r2 As Range, r3 As Range, r4 As Range)
    ' 複数列バリューに対応させた
     Dim myDic As Object, myKey As Variant
     Dim c As Variant

     Dim myrows As Long
     Dim mycolumns As Long

     Dim myRange As Range

     On Error GoTo KeyValueSet_Error

     Set myRange = Range(r1, r1.End(xlDown))

     myrows = myRange.Rows.Count
     mycolumns = myRange.columns.Count

     If mycolumns <> 1 Then
         MsgBox "選択列を１列にしてください。"
         GoTo KeyValueSet_Normal_exit
     End If

     Dim valueRange As Range
     Set valueRange = r2
     Dim valuecols As Long
     valuecols = valueRange.columns.Count

     Dim distKeyRange As Range
     Set distKeyRange = r3
     Set distKeyRange = Range(distKeyRange, distKeyRange.End(xlDown))

     Dim distValueRange As Range
     Set distValueRange = r4
     
     Dim i As Long
     Dim j As Long
     Dim varData As Variant
     Dim varValue As Variant
     Dim itemData As Variant
     Set myDic = CreateObject("Scripting.Dictionary")
     varData = myRange.Value
     varValue = valueRange.Resize(myrows, valuecols).Value

     Dim vararray() As Variant
     For i = 1 To myrows
         c = varData(i, 1)
         If Not c = Empty Then
             If Not myDic.Exists(c) Then
                 ReDim vararray(1 To valuecols)
                 For j = 1 To valuecols
                     vararray(j) = varValue(i, j)
                 Next j
                 myDic.Add CStr(c), vararray
             Else
                 MsgBox "キーが重複しています::" & c
                 GoTo KeyValueSet_Normal_exit
             End If
         End If
     Next


     Dim itemval As String
     Dim keytemp As String
     If WorksheetFunction.CountA(distValueRange.Resize(distKeyRange.Rows.Count, valuecols)) <> 0 Then
         MsgBox "出力先が空欄でありません"
         GoTo KeyValueSet_Normal_exit
     Else
         For i = 1 To distKeyRange.Rows.Count
             keytemp = CStr(distKeyRange.Cells(i, 1))
             If myDic.Exists(keytemp) Then
                 For j = 1 To valuecols
                     distValueRange.Resize(myrows, valuecols).Cells(i, j) = myDic.Item(keytemp)(j)
                 Next j
             Else
                 For j = 1 To valuecols
                     distValueRange.Resize(myrows, valuecols).Cells(i, j) = "#no val#"
                 Next j
             End If

         Next i
     End If



     On Error GoTo 0
     GoTo KeyValueSet_Normal_exit

KeyValueSet_Error:
     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure KeyValueSet2 of Module modIROIRO"

KeyValueSet_Normal_exit:
     Set myDic = Nothing

End Sub



Sub KeyValueSetMatrix()
    Dim myDic As Object, myKey As Variant
    Dim c As Variant
    
    Dim myrows As Long
    Dim mycolumns As Long
    
    Dim myRange As Range
    
    'On Error GoTo KeyValueSetMatrix_Error

    Set myRange = Range(Selection.Address)
    
    myrows = myRange.Rows.Count
    mycolumns = myRange.columns.Count
    
    
    If mycolumns <> 2 Then
        MsgBox "選択列を2列にしてください。"
        GoTo KeyValueSetMatrix_Normal_exit
    End If
    
    
    Dim distRange As Range
    Set distRange = Application.InputBox(prompt:="", title:="適用データ", Type:=8)
    
    Dim distCopy As Range
    Set distCopy = Application.InputBox(prompt:="", title:="適用後コピー先", Type:=8)
    
    
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.StatusBar = ""
    
    Dim i As Long
    Dim varData As Variant
    Dim itemData As Variant
    Set myDic = CreateObject("Scripting.Dictionary")
    varData = myRange.Value
   
    For i = 1 To myrows
        c = CStr(varData(i, 1))
        If Not c = Empty Then
            If Not myDic.Exists(c) Then
                myDic.Add c, CStr(varData(i, 2))
            Else
                MsgBox "キーが重複しています::" & c
                GoTo KeyValueSetMatrix_Normal_exit
            End If
        End If
    Next
    
    '配列が返ってくる
    Dim j As Long
    Dim k As Long
    Dim itemval As String
    Dim keytemp As String
    Dim offsetcol As Long
    

    If WorksheetFunction.CountA(distCopy.Resize(distRange.Rows.Count, distRange.columns.Count)) <> 0 Then
        MsgBox "出力先が空欄でありません" & distCopy.Address
        GoTo KeyValueSetMatrix_Normal_exit
    Else
        Dim varDataDist() As Variant
        ReDim varDataDist(1 To distRange.Rows.Count, 1 To distRange.columns.Count)
        
        For j = 1 To distRange.Rows.Count
            For k = 1 To distRange.columns.Count
                keytemp = CStr(distRange.Cells(j, k))
                If myDic.Exists(keytemp) Then
                    itemval = myDic.Item(keytemp)
                ElseIf keytemp <> "" Then
                    itemval = "###### No Item #####"
                Else
                    itemval = ""
                End If
                'distCopy.Offset(j - 1, k - 1).Cells(1, 1) = itemval
                varDataDist(j, k) = UCase(Left(itemval, 1)) & Mid(itemval, 2)
            Next k
        Next j
    End If
    
    distCopy.Resize(distRange.Rows.Count, distRange.columns.Count) = varDataDist

    On Error GoTo 0
    GoTo KeyValueSetMatrix_Normal_exit

KeyValueSetMatrix_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure KeyValueSetMatrix of Module modIROIRO"

KeyValueSetMatrix_Normal_exit:
    Set myDic = Nothing
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
End Sub

Public Sub SaveRngToFile(ByRef rngForSave As Range, ByVal Fname As String)
    
    Dim varData As Variant
    varData = rngForSave.Value
    
    
    Dim DataList As clsString
    Set DataList = New clsString

    Dim DataLine As clsString

    Dim i As Long
    Dim j As Long
    Dim columns As Long

    Dim tmpstr As String
    Dim isquate As Boolean
    
    
    Dim myrow As Long
    Dim mycol As Long
    myrow = rngForSave.Rows.Count
    mycol = rngForSave.columns.Count


    For i = 1 To myrow
        Set DataLine = New clsString
        For j = 1 To mycol
            If mycol = 1 And myrow = 1 Then
                 tmpstr = CStr(varData)
            Else
                tmpstr = CStr(varData(i, j))
            End If
           
'            '' , " のエスケープ
'            If InStr(1, tmpstr, """") Or InStr(1, tmpstr, ",") Then
'                tmpstr = """" & Replace(tmpstr, """", """""") & """"
'            End If

            DataLine.Add tmpstr

        Next j
        DataList.Add DataLine.Joins(vbTab)
    Next i

    DataList.SaveToFileUTF8 Fname
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DiffRange
' Author    : toramame
' Date      : 2016/12/21
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub DiffRange()
    On Error GoTo DiffRange_Error
    
    Dim outputpath As String
    outputpath = "C:\temp"
    
    If FsoFolderExists(outputpath) = False Then
        MsgBox outputpath & "を作成してください。"
        Exit Sub
    End If
    
    
    Dim rng1 As Range
    Set rng1 = Range(Selection.Address) ''Application.InputBox(Prompt:="", title:="領域1", Type:=8)
    
    Dim rng2 As Range
    Set rng2 = Application.InputBox(prompt:="", title:="比較対象のはじめのセル", Type:=8)
    Set rng2 = rng2.Resize(rng1.Rows.Count, rng1.columns.Count)
    
    
    Dim FileName1 As String
    Dim FileName2 As String
    FileName1 = outputpath & "\" & Format(Now, "MMDD-HHmmSS") & "-はじめ.txt"
    FileName2 = outputpath & "\" & Format(Now, "MMDD-HHmmSS") & "-あと.txt"
    
    SaveRngToFile rng1, FileName1
    SaveRngToFile rng2, FileName2
    
    Dim strcmds As String
    strcmds = "'C:\Program Files\WinMerge\WinMergeU.exe' " & "'" & FileName1 & "' '" & FileName2 & "'"
    

    Shell Replace(strcmds, "'", """")

    On Error GoTo 0
    GoTo DiffRange_Normal_exit

DiffRange_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DiffRange of Module modIROIRO"

DiffRange_Normal_exit:

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ReplaceWithList
' Author    : toramame
' Date      : 2017/01/13
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub ReplaceWithList()
    Dim rng1 As Range
    Dim rng2 As Range

    'On Error GoTo ReplaceWithList_Error
    
    Dim rngSelection As Range
    

    Set rngSelection = Selection

    Set rng1 = Range(Selection.Address) ''Application.InputBox(Prompt:="", title:="領域1", Type:=8)
    
    Set rng2 = Application.InputBox(prompt:="置換対象を選択してから、リスト2列を選ぶ", title:="検索対象、置換文字", Type:=8)
    
    If rng2.columns.Count <> 2 Then
        MsgBox "置換リストは2列選んでください。"
        Exit Sub
    End If
    
    If MsgBox("一括置換しますがよろしいですか?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Dim i As Long
    Dim j As Long
    
    Dim len1 As Long
    Dim len2 As Long
    Dim text1 As String
    Dim text2 As String
    Dim textSelection As String
    
        For i = 1 To rng2.Rows.Count
            text1 = rng2.Cells(i, 1)
            text2 = rng2.Cells(i, 2)
            
            len1 = Len(text1)
            len1 = Len(text2)
            
            If (len1 > 255 Or len2 > 255) Then
                For j = 1 To rngSelection.Rows.Count
                    textSelection = rngSelection.Cells(j, 1).text
                    rngSelection.Cells(j, 1) = Replace(textSelection, text1, text2, 1, -1, vbTextCompare)
                Next j
            Else
                rngSelection.Replace What:=rng2.Cells(i, 1), Replacement:=rng2.Cells(i, 2), _
                LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False, Matchbyte:=True
            End If
            
            Application.StatusBar = i / rng2.Rows.Count
            DoEvents
        Next i


    On Error GoTo 0
    GoTo ReplaceWithList_Normal_exit
    
    
ReplaceWithList_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ReplaceWithList of Module modIROIRO"

ReplaceWithList_Normal_exit:

   On Error GoTo 0

    Application.StatusBar = False
       Exit Sub
End Sub

Function GetPacID(s As String) As String
'レジストリのIDを調べる

Dim ss As String
ss = Replace(s, "-", "")

Dim s1(0 To 31) As String
Dim s2(0 To 31) As String

Dim i As Long
For i = 0 To 31
    s1(i) = Mid(ss, i + 1, 1)
Next i


s2(0) = s1(7)
s2(1) = s1(6)
s2(2) = s1(5)
s2(3) = s1(4)
s2(4) = s1(3)
s2(5) = s1(2)
s2(6) = s1(1)
s2(7) = s1(0)
s2(8) = s1(11)
s2(9) = s1(10)
s2(10) = s1(9)
s2(11) = s1(8)
s2(12) = s1(15)
s2(13) = s1(14)
s2(14) = s1(13)
s2(15) = s1(12)
s2(16) = s1(17)
s2(17) = s1(16)
s2(18) = s1(19)
s2(19) = s1(18)
s2(20) = s1(21)
s2(21) = s1(20)
s2(22) = s1(23)
s2(23) = s1(22)
s2(24) = s1(25)
s2(25) = s1(24)
s2(26) = s1(27)
s2(27) = s1(26)
s2(28) = s1(29)
s2(29) = s1(28)
s2(30) = s1(31)
s2(31) = s1(30)

GetPacID = Join(s2, "")

End Function

Sub Diff2Lines()
    Dim range1 As Range, range2 As Range, rangeDiff As Range
       
    Set range1 = Selection
    If range1.columns.Count > 1 Then
        MsgBox "2列以上は選択できません"
        Exit Sub
    End If
    
    Set range2 = Application.InputBox(prompt:="", title:="2列目を指定する", Type:=8)
    If range2.columns.Count > 1 Then
        MsgBox "2列以上は選択できません"
        Exit Sub
    End If
    
    Set rangeDiff = Application.InputBox(prompt:="", title:="出力先を指定する", Type:=8)
    
    If WorksheetFunction.CountA(rangeDiff.Resize(range1.Rows.Count + range2.Rows.Count, 3)) <> 0 Then
        MsgBox "比較結果出力先が空欄でありません"
        Exit Sub
    End If

    Dim r(1 To 2) As Variant
    r(1) = range1
    r(2) = range2
       
    Dim dic(1 To 2) As Object, dicDiff As Object
    Set dic(1) = CreateObject("Scripting.Dictionary")
    Set dic(2) = CreateObject("Scripting.Dictionary")
    Set dicDiff = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim c As Long
    Dim v As String
    For c = 1 To 2
      For i = 1 To UBound(r(c))
          v = r(c)(i, 1)
          If Not dicDiff.Exists(v) Then
              dicDiff.Add v, v
          End If
          If Not dic(c).Exists(v) Then
            dic(c).Add v, v
          End If
      Next
    Next

    Dim allKeys As Variant  '
    allKeys = dicDiff.Keys
    
    Dim diffArrayList As Object
    Set diffArrayList = CreateObject("System.Collections.ArrayList")
    
    For i = 0 To UBound(allKeys)
        diffArrayList.Add (allKeys(i))
    Next
    diffArrayList.Sort
    
    Dim outValue As Variant
    outValue = rangeDiff.Resize(UBound(allKeys) + 1, 3)
    For i = 0 To UBound(allKeys)
        outValue(i + 1, 1) = diffArrayList(i)
        outValue(i + 1, 2) = IIf(dic(1).Exists(diffArrayList(i)), 1, 0)
        outValue(i + 1, 3) = IIf(dic(2).Exists(diffArrayList(i)), 1, 0)
    Next
    rangeDiff.Resize(UBound(allKeys) + 1, 3) = outValue
End Sub

