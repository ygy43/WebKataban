Imports WebKataban.ClsCommon
Imports System.Data.SqlClient
Imports WebKataban.CdCst

Public Class SiyouBLL

#Region "形番分解"
    ''' <summary>
    ''' フル形番よりキー形番を分解する
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="HTSelKata"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSelKata(ByVal objCon As SqlConnection, ByVal objKtbnStrc As KHKtbnStrc, _
                                      ByRef HTSelKata As ArrayList) As Boolean
        Dim strPriceDiv() As String = Nothing
        GetSelKata = False
        Try
            Dim strSpecNo As String = objKtbnStrc.strcSelection.strSpecNo
            Dim strValue As New ArrayList
            Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban
            Dim strKeyKata As String = objKtbnStrc.strcSelection.strKeyKataban

            strValue.Add(objKtbnStrc.strcSelection.strSeriesKataban)
            For inti As Integer = 1 To objKtbnStrc.strcSelection.strOpSymbol.Length - 1
                strValue.Add(objKtbnStrc.strcSelection.strOpSymbol(inti))
            Next

            If strSpecNo.Trim.Length <= 0 Then Exit Function
            HTSelKata = New ArrayList
            For inti As Integer = 0 To 40
                Dim listC As New ArrayList
                HTSelKata.Add(listC)
            Next

            Dim dt As DataTable = LoadComboData(objCon, strSpecNo.Trim)

            Dim intMaxCount As Integer = 0
            If strValue.Count > 0 Then
                intMaxCount = strValue.Count - 1
            End If

            If dt.Rows.Count > 0 Then
                For inti As Integer = 0 To dt.Rows.Count - 1
                    'シリアル形番を取得する
                    Dim dr As DataRow = dt.Rows(inti)
                    Dim strLine() As String = dr("LineNo").ToString.Split(",")
                    Dim strSeries() As String = dr("SeriesKata").ToString.Split(CdCst.Sign.Delimiter.Comma)
                    For intk As Integer = 0 To strSeries.Length - 1
                        '同じシリアル形番+同じキー形番
                        Dim strKey() As String = dr("KeyKata").ToString.Split(CdCst.Sign.Delimiter.Comma)
                        For intM As Integer = 0 To strKey.Length - 1
                            If strSeriesKata.Trim = strSeries(intk) And _
                                (strKey(intM).ToString = "ALL" Or _
                                 strKey(intM).ToString = strKeyKata) Then
                                '条件ﾁｪｯｸ
                                If CheckWhere(dt, dr, strValue, strSpecNo, strSeriesKata, strKeyKata, intMaxCount) Then
                                    Dim ListGroup As New ArrayList
                                    ListGroup = CreatSelKata(strSeriesKata, dt, dr, strValue)
                                    For intj As Integer = 0 To strLine.Length - 1
                                        If strLine(intj).ToString.Length <= 0 Then Continue For
                                        Dim lst As New ArrayList
                                        For intl As Integer = 0 To ListGroup.Count - 1
                                            Select Case dr("ColNo")
                                                Case 0               '形番
                                                    If Not HTSelKata(CInt(strLine(intj).ToString)) Is Nothing AndAlso _
                                                       HTSelKata(CInt(strLine(intj).ToString)).Count > 0 Then
                                                        lst = HTSelKata(CInt(strLine(intj).ToString))
                                                    End If
                                            End Select
                                            lst.Add(ListGroup(intl))
                                        Next
                                        If dr("Empty") And lst.Count > 0 Then
                                            If lst(0) <> String.Empty Then
                                                lst.Insert(0, String.Empty)
                                            End If
                                        End If
                                        Select Case dr("ColNo")
                                            Case 0               '形番
                                                HTSelKata(CInt(strLine(intj).ToString)) = lst
                                        End Select
                                    Next
                                End If
                            ElseIf strSeries(intk) = "ALL" And (strKey(intM).ToString = "ALL" Or _
                                strKey(intM).ToString = strKeyKata) Then            '全シリーズ
                                '条件ﾁｪｯｸ
                                If CheckWhere(dt, dr, strValue, strSpecNo, strSeriesKata, strKeyKata, intMaxCount) Then
                                    Dim ListGroup As New ArrayList
                                    ListGroup = CreatSelKata(strSeriesKata, dt, dr, strValue)
                                    For intj As Integer = 0 To strLine.Length - 1
                                        If strLine(intj).ToString.Length <= 0 Then Continue For
                                        Dim lst As New ArrayList
                                        For intl As Integer = 0 To ListGroup.Count - 1
                                            Select Case dr("ColNo")
                                                Case 0               '形番
                                                    If Not HTSelKata(CInt(strLine(intj).ToString)) Is Nothing AndAlso _
                                                       HTSelKata(CInt(strLine(intj).ToString)).Count > 0 Then
                                                        lst = HTSelKata(CInt(strLine(intj).ToString))
                                                    End If
                                            End Select
                                            lst.Add(ListGroup(intl))
                                        Next
                                        If dr("Empty") And lst.Count > 0 Then
                                            If lst(0) <> String.Empty Then
                                                lst.Insert(0, String.Empty)
                                            End If
                                        End If
                                        Select Case dr("ColNo")
                                            Case 0               '形番
                                                HTSelKata(CInt(strLine(intj).ToString)) = lst
                                        End Select
                                    Next
                                End If
                            End If
                        Next
                    Next
                Next
            End If

            GetSelKata = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 条件ﾁｪｯｸ
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="dr"></param>
    ''' <param name="dr_value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckWhere(ByVal dt As DataTable, ByVal dr As DataRow, ByVal dr_value As ArrayList, _
                                       ByVal strSpecNo As String, ByVal strSeriesKata As String, _
                                       ByVal strKeyKata As String, ByVal intMaxCount As Integer) As Boolean
        CheckWhere = False
        Try
            '条件解析
            For intL As Integer = 0 To dt.Columns.Count - 1
                If dt.Columns(intL).ColumnName Like "Item*" Then
                    Dim intCount As Integer = CInt(Strings.Right(dt.Columns(intL).ColumnName, 1))
                    If dr("Item" & intCount).ToString.Length > 0 Then
                        Dim strVal As String = dr("Value" & intCount).ToString
                        If strVal = "<>!" Then           '空白不可
                            If dr_value(dr("Item" & intCount)).ToString.Trim.Length <= 0 Then
                                Return False
                            Else
                                Continue For                       '次の条件をチェック
                            End If
                        ElseIf strVal = "=!" Then        '空白
                            If dr_value(dr("Item" & intCount)).ToString.Trim.Length > 0 Then
                                Return False
                            Else
                                Continue For                       '次の条件をチェック
                            End If
                        ElseIf strVal.StartsWith("[SPEC]=!") Then   '仕様書Noの判断
                            If strVal.Equals(strSpecNo) Then
                                Continue For                       '次の条件をチェック
                            Else
                                Return False
                            End If
                        ElseIf strVal.StartsWith("[SPEC]<>!") Then   '仕様書Noの判断
                            If Not strVal.Equals(strSpecNo) Then
                                Continue For                       '次の条件をチェック
                            Else
                                Return False
                            End If
                        ElseIf strVal.StartsWith("[SPEC]=") Then   '仕様書Noの判断
                            Dim str() As String = strVal.Split("=")
                            If str.Length = 2 Then
                                If str(1).EndsWith(strSpecNo) Then
                                    Continue For                       '次の条件をチェック
                                Else
                                    Return False
                                End If
                            Else
                                Return False
                            End If
                        ElseIf strVal.StartsWith("[SPEC]<>") Then   '仕様書Noの判断
                            Dim str() As String = strVal.Split(">")
                            If str.Length = 2 Then
                                If Not str(1).EndsWith(strSpecNo) Then
                                    Continue For                       '次の条件をチェック
                                Else
                                    Return False
                                End If
                            Else
                                Return False
                            End If
                        ElseIf strVal.StartsWith("START=") Then    '頭
                            Dim str() As String = dr_value(dr("Item" & intCount)).ToString.Split(",")
                            Dim bol As Boolean = False
                            For inti As Integer = 0 To str.Length - 1
                                If str(inti).ToString.Length <= 0 Then
                                    Continue For
                                Else
                                    If Not str(inti).ToString.StartsWith(strVal.Split("=")(1)) Then
                                        Continue For
                                    Else
                                        bol = True
                                        Exit For                       '次の条件をチェック
                                    End If
                                End If
                            Next
                            If Not bol Then
                                Return False
                            End If
                        ElseIf strVal.StartsWith("START<>") Then    '頭
                            Dim str() As String = dr_value(dr("Item" & intCount)).ToString.Split(",")
                            Dim bol As Boolean = False
                            For inti As Integer = 0 To str.Length - 1
                                If str(inti).ToString.Length <= 0 Then
                                    Continue For
                                Else
                                    If str(inti).ToString.StartsWith(strVal.Split(">")(1)) Then
                                        Return False
                                    Else
                                        Continue For                       '次の条件をチェック
                                    End If
                                End If
                            Next
                        ElseIf strVal.StartsWith("END=") Then    '頭
                            Dim str() As String = strVal.Split("=")(1).Split(",")
                            Dim bol As Boolean = False
                            For inti As Integer = 0 To str.Length - 1
                                If dr_value(dr("Item" & intCount)).ToString.EndsWith(str(inti)) Then
                                    bol = True
                                    Exit For
                                End If
                            Next
                            If bol Then
                                Continue For                       '次の条件をチェック
                            Else
                                Return False
                            End If
                        ElseIf strVal.StartsWith("END<>") Then    '頭
                            Dim str() As String = strVal.Split(">")(1).Split(",")
                            Dim bol As Boolean = False
                            For inti As Integer = 0 To str.Length - 1
                                If dr_value(dr("Item" & intCount)).ToString.EndsWith(str(inti)) Then
                                    Return False
                                End If
                            Next
                        ElseIf strVal.StartsWith("<>NOTHING") Then    '頭
                            If intMaxCount >= CInt(dr("Item" & intCount).ToString) Then
                                Continue For                       '次の条件をチェック
                            Else
                                Return False
                            End If
                        ElseIf strVal.StartsWith("=NOTHING") Then    '頭
                            If intMaxCount >= CInt(dr("Item" & intCount).ToString) Then
                                Return False                       '次の条件をチェック
                            Else
                                Continue For
                            End If
                        ElseIf strVal.StartsWith("MID(") Then    '頭
                            Dim intlen As Integer = 0
                            Dim str() As String = strVal.Split(",")
                            If str.Length = 2 Then
                                intlen = CInt(Strings.Right(str(0), 1)) + CInt(Strings.Left(str(1), 1))
                            End If
                            If dr_value(dr("Item" & intCount)).ToString.Length < intlen Then
                                Return False
                            End If
                            Dim str_1() As String = strVal.Split("=")
                            If str_1.Length = 2 Then
                                If Mid(dr_value(dr("Item" & intCount)).ToString, CInt(str(0)), CInt(str(1))) = str_1(1) Then
                                    Continue For
                                Else
                                    Return False
                                End If
                            End If
                        End If
                        If strVal.StartsWith("=") Then             '=条件
                            Dim strchild() As String = dr_value(dr("Item" & intCount)).ToString.Trim.Split(",")
                            Dim chk As Boolean = False
                            Dim strKey() As String = Strings.Right(strVal, strVal.Length - 1).ToString.Trim.Split(",")
                            For intj As Integer = 0 To strKey.Length - 1
                                If strKey(intj) = "!" Then
                                    strKey(intj) = String.Empty
                                End If
                                For inti As Integer = 0 To strchild.Length - 1
                                    If strchild(inti) = strKey(intj) Then
                                        chk = True
                                        Exit For
                                    End If
                                Next
                                If chk Then Exit For
                            Next
                            If Not chk Then
                                Return False
                            Else
                                Continue For
                            End If
                        ElseIf strVal.StartsWith("<>") Then             '<>条件
                            Dim strchild() As String = dr_value(dr("Item" & intCount)).ToString.Trim.Split(",")
                            Dim chk As Boolean = False
                            Dim strKey() As String = Strings.Right(strVal, strVal.Length - 2).ToString.Trim.Split(",")
                            For intj As Integer = 0 To strKey.Length - 1
                                If strKey(intj) = "!" Then
                                    If dr_value(dr("Item" & intCount)).ToString.Trim.Length = 0 Then
                                        Return False
                                    End If
                                End If
                                For inti As Integer = 0 To strchild.Length - 1
                                    If strchild(inti) = strKey(intj) Then
                                        Return False
                                    End If
                                Next
                            Next
                        End If
                    Else
                        Exit For
                    End If
                End If
            Next
            CheckWhere = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 価格計算用のキー形番を生成する
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="dr"></param>
    ''' <param name="dr_value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CreatSelKata(ByVal strSeriesKata As String, ByVal dt As DataTable, _
                                         ByVal dr As DataRow, ByVal dr_value As ArrayList) As ArrayList
        CreatSelKata = New ArrayList
        Try
            '価格計算用のキー形番を生成する
            For intL As Integer = 0 To dt.Columns.Count - 1
                If dt.Columns(intL).ColumnName Like "KeyValue*" Then
                    Dim intCount As Integer = CInt(Strings.Right(dt.Columns(intL).ColumnName, 1))
                    If dr("KeyValue" & intCount).ToString.Length > 0 Then     '空白ではない場合
                        Dim strVal As String = dr("KeyValue" & intCount).ToString
                        If strVal = "[0]" Then                                'シリアル形番
                            If CreatSelKata.Count = 0 Then
                                CreatSelKata.Add(strSeriesKata)
                            Else
                                For inti As Integer = 0 To CreatSelKata.Count - 1
                                    CreatSelKata.Item(inti) &= strSeriesKata
                                Next
                            End If
                        ElseIf strVal.StartsWith("LEFT") Then                 '頭何桁目のみ 
                            Dim str() As String = strVal.Split("-")
                            If str.Length = 2 Then
                                If str(1).StartsWith("[") And str(1).EndsWith("]") Then
                                    Dim str1 As String = dr_value(Mid(str(1), 2, str(1).Length - 2)).ToString.Trim
                                    If str1.Length >= CInt(Mid(str(0), 5, str(0).Length - 4)) Then
                                        If CreatSelKata.Count = 0 Then
                                            CreatSelKata.Add(Left(str1, CInt(Mid(str(0), 5, str(0).Length - 4))))
                                        Else
                                            For inti As Integer = 0 To CreatSelKata.Count - 1
                                                CreatSelKata.Item(inti) &= Left(str1, CInt(Mid(str(0), 5, str(0).Length - 4)))
                                            Next
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf strVal.StartsWith("RIGHT") Then                 '最後何桁目のみ 
                            Dim str() As String = strVal.Split("-")
                            If str.Length = 2 Then
                                If str(1).StartsWith("[") And str(1).EndsWith("]") Then
                                    Dim str1 As String = dr_value(Mid(str(1), 2, str(1).Length - 2)).ToString.Trim
                                    If str1.Length >= CInt(Mid(str(0), 6, str(0).Length - 5)) Then
                                        If CreatSelKata.Count = 0 Then
                                            CreatSelKata.Add(Right(str1, CInt(Mid(str(0), 6, str(0).Length - 5))))
                                        Else
                                            For inti As Integer = 0 To CreatSelKata.Count - 1
                                                CreatSelKata.Item(inti) &= Right(str1, CInt(Mid(str(0), 6, str(0).Length - 5)))
                                            Next
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf strVal.StartsWith("MID(") Then
                            Dim intlen As Integer = 0
                            Dim str() As String = strVal.Split("-")
                            If str.Length = 3 Then
                                intlen = CInt(str(1)) + CInt(Strings.Left(str(2), 1))
                            End If
                            Dim str1 As String = dr_value(Mid(str(0), 6, 1)).ToString
                            If str1.Length >= intlen Then
                                If CreatSelKata.Count = 0 Then
                                    CreatSelKata.Add(Mid(str1, CInt(str(1)), CInt(Strings.Left(str(2), 1))))
                                Else
                                    For inti As Integer = 0 To CreatSelKata.Count - 1
                                        CreatSelKata.Item(inti) &= Mid(str1, CInt(str(1)), CInt(Strings.Left(str(2), 1)))
                                    Next
                                End If
                            End If
                        ElseIf strVal.StartsWith("[") And strVal.EndsWith("]") Then     '複数なら、形番行を追加
                            Dim strOption() As String = dr_value(Mid(strVal, 2, strVal.Length - 2)).ToString.Split(CdCst.Sign.Delimiter.Comma)
                            If strOption.Length <= 1 Then
                                If CreatSelKata.Count = 0 Then
                                    CreatSelKata.Add(dr_value(Mid(strVal, 2, strVal.Length - 2)).ToString)
                                Else
                                    For inti As Integer = 0 To CreatSelKata.Count - 1
                                        CreatSelKata.Item(inti) &= dr_value(Mid(strVal, 2, strVal.Length - 2)).ToString
                                    Next
                                End If
                            Else
                                If CreatSelKata.Count = 0 Then
                                    For inti As Integer = 0 To strOption.Length - 1
                                        If strOption(inti).ToString.Trim.Length <= 0 Then Continue For
                                        CreatSelKata.Add(strOption(inti).ToString.Trim)
                                    Next
                                Else
                                    Dim newlist As New ArrayList
                                    For inti As Integer = 0 To CreatSelKata.Count - 1
                                        For intj As Integer = 0 To strOption.Length - 1
                                            If strOption(intj).ToString.Trim.Length <= 0 Then Continue For
                                            newlist.Add(CreatSelKata.Item(inti) & strOption(intj))
                                        Next
                                    Next
                                    If newlist.Count > 0 Then
                                        CreatSelKata = newlist
                                    End If
                                End If
                            End If
                        ElseIf strVal.StartsWith("(") And strVal.EndsWith(")") Then      '複数なら、形番行を追加しない
                            Dim strOption() As String = dr_value(Mid(strVal, 2, strVal.Length - 2)).ToString.Split(CdCst.Sign.Delimiter.Comma)
                            If CreatSelKata.Count = 0 Then
                                CreatSelKata.Add("")
                            End If
                            For inti As Integer = 0 To CreatSelKata.Count - 1
                                For intj As Integer = 0 To strOption.Length - 1
                                    If strOption(intj).ToString.Trim.Length <= 0 Then Continue For
                                    CreatSelKata.Item(inti) &= strOption(intj).ToString.Trim
                                Next
                            Next
                        ElseIf strVal.StartsWith("<") And strVal.EndsWith(">") Then     'キー値のみ
                            Dim intSeq As Integer = CInt(Mid(strVal, 2, strVal.Length - 2))
                            Dim intKey As Integer = dr("Item" & intSeq).ToString
                            Dim strkey As String = dr("Value" & intSeq).ToString
                            Dim strOption() As String = dr_value(intKey).Split(CdCst.Sign.Delimiter.Comma)
                            Dim strKeyOption() As String = Strings.Right(strkey, strkey.Length - 1).Split(CdCst.Sign.Delimiter.Comma)

                            If CreatSelKata.Count = 0 Then
                                For inti As Integer = 0 To strOption.Length - 1
                                    If strOption(inti).ToString.Length <= 0 Then Continue For
                                    For intj As Integer = 0 To strKeyOption.Length - 1
                                        If strOption(inti) = strKeyOption(intj) Then
                                            CreatSelKata.Add(strOption(inti))
                                        End If
                                    Next
                                Next
                            Else
                                For intk As Integer = 0 To CreatSelKata.Count - 1
                                    For inti As Integer = 0 To strOption.Length - 1
                                        If strOption(inti).ToString.Length <= 0 Then Continue For
                                        For intj As Integer = 0 To strKeyOption.Length - 1
                                            If strOption(inti) = strKeyOption(intj) Then
                                                CreatSelKata.Item(intk) &= strOption(inti)
                                            End If
                                        Next
                                    Next
                                Next
                            End If
                        Else
                            Dim str() As String = strVal.Split(",")
                            For inti As Integer = 0 To str.Length - 1
                                If str(inti).ToString.Length <= 0 Then Continue For
                                If str(inti) = "{REPJ}" Then  '検査成績書
                                    str(inti) = CdCst.Manifold.InspReportJp.Japanese
                                End If
                                If str(inti) = "{REPE}" Then
                                    str(inti) = CdCst.Manifold.InspReportEn.Japanese
                                End If
                            Next
                            If CreatSelKata.Count = 0 Then
                                For inti As Integer = 0 To str.Length - 1
                                    CreatSelKata.Add(str(inti))
                                Next
                            Else
                                Dim newlist As New ArrayList
                                For inti As Integer = 0 To CreatSelKata.Count - 1
                                    For intj As Integer = 0 To str.Length - 1
                                        newlist.Add(CreatSelKata.Item(inti) & str(intj))
                                    Next
                                Next
                                If newlist.Count > 0 Then
                                    CreatSelKata = newlist
                                End If
                            End If
                        End If

                        Dim bolHypen As Boolean = False
                        If dt.Columns("KeyHypen" & intCount) Is Nothing Then
                            bolHypen = False
                        Else
                            bolHypen = dr("KeyHypen" & intCount)
                        End If

                        If bolHypen Then
                            For inti As Integer = 0 To CreatSelKata.Count - 1
                                CreatSelKata.Item(inti) &= "-"
                            Next
                        End If
                    Else
                        Exit For
                    End If
                End If
            Next
            For inti As Integer = 0 To CreatSelKata.Count - 1
                If CreatSelKata.Item(inti).ToString.EndsWith("----") Then
                    CreatSelKata.Item(inti) = Left(CreatSelKata.Item(inti).ToString, CreatSelKata.Item(inti).ToString.Length - 4)
                ElseIf CreatSelKata.Item(inti).ToString.EndsWith("---") Then
                    CreatSelKata.Item(inti) = Left(CreatSelKata.Item(inti).ToString, CreatSelKata.Item(inti).ToString.Length - 3)
                ElseIf CreatSelKata.Item(inti).ToString.EndsWith("--") Then
                    CreatSelKata.Item(inti) = Left(CreatSelKata.Item(inti).ToString, CreatSelKata.Item(inti).ToString.Length - 2)
                ElseIf CreatSelKata.Item(inti).ToString.EndsWith("-") Then
                    CreatSelKata.Item(inti) = Left(CreatSelKata.Item(inti).ToString, CreatSelKata.Item(inti).ToString.Length - 1)
                End If
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' マニホールド画面仕様を取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSpecNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function LoadComboData(ByVal objCon As SqlConnection, strSpecNo As String) As DataTable
        Dim dtResult As New DataTable
        Dim dalSiyou As New SiyouDAL

        Try
            dtResult = dalSiyou.LoadComboData(objCon, strSpecNo)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return dtResult
    End Function
#End Region

#Region "取付レールの計算"
    ''' <summary>
    ''' 取付レール値計算・取得
    ''' </summary>
    ''' <param name="ds">選択情報</param>
    ''' <param name="ManifoldMode"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intRailRowID">レール長さ行番号</param>
    ''' <param name="strRailChangeFlg">レール長さ変更フラグ</param>
    ''' <param name="dblRailLen">レール長さ</param>
    ''' <param name="dblStdNum">レール標準</param>
    ''' <remarks></remarks>
    Public Sub subGetRail(ds As DataSet, ByVal ManifoldMode As Integer, ByRef objKtbnStrc As KHKtbnStrc, _
                           ByVal intRailRowID As Integer, ByVal strRailChangeFlg As String, _
                           ByRef dblRailLen As Double, ByRef dblStdNum As Double)

        Dim intPartCnt As Integer = 0
        Dim intValveCnt1 As Integer = 0
        Dim intValveCnt2 As Integer = 0
        Dim intExhaustCnt1 As Integer = 0
        Dim intExhaustCnt2 As Integer = 0
        Dim dblR1 As Double = 0
        Dim dblX As Double

        Dim intValveCnt As Integer = 0
        Dim intExhaustCnt As Integer = 0
        Dim intEndLCnt As Integer = 0
        Dim intEndRCnt As Integer = 0
        Dim intReguCnt As Integer = 0
        Dim intDummyCnt As Integer = 0
        Dim intTLCnt As Integer = 0
        Dim intTMCnt As Integer = 0
        Dim intTRCnt As Integer = 0
        Dim intT6Cnt As Integer = 0
        Dim intT7Cnt As Integer = 0
        Dim intValveWidth As Integer        'バルブブロック幅
        Dim intValveCnt07 As Integer        '7mmバルブ使用数
        Dim intValveCnt10 As Integer        '10mmバルブ使用数
        Dim intT7ECCnt As Integer = 0 '2016/08/23 RM1608024 K.Ohwaki Append

        Dim intPositionCnt As Integer = 0
        Dim dblY As Double
        Dim dblZ As Double

        Dim strInpUse As New ArrayList
        Dim strSelKataban As New ArrayList

        '手動入力レール長さ
        Dim intSelRail As Decimal

        '選択した形番と使用数の取得
        subGetInfoFromDS(ds, strSelKataban, strInpUse, ManifoldMode)

        '手動入力レール長さ
        intSelRail = IIf(strInpUse(intRailRowID).Equals(String.Empty), 0, CDec(strInpUse(intRailRowID)))

        'レール長さの計算
        Try
            Select Case ManifoldMode
                Case 1
                    'シリーズでバルブ幅を判別
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "MN3E00", "MN3EX0", "MN4E00", "MN4EX0" '7mmバルブ
                            intValveWidth = 7   '7mm
                        Case Else                                   '10mmバルブ
                            intValveWidth = 10  '10mm
                    End Select

                    If strInpUse Is Nothing Then
                    ElseIf strInpUse.Count = 0 Then
                    Else
                        '使用数カウント
                        For intI As Integer = CdCst.Siyou_01.Elect1 - 1 To CdCst.Siyou_01.Elect2 - 1
                            If Int(strInpUse(intI)) > 0 Then
                                If Len(strSelKataban(intI)) < 7 Then
                                ElseIf strSelKataban(intI).Substring(5, 2) = "TM" Then
                                    intTMCnt = intTMCnt + Int(strInpUse(intI))

                                ElseIf strSelKataban(intI).Substring(5, 2) = "T3" Or _
                                        strSelKataban(intI).Substring(5, 2) = "T5" Then

                                    If strSelKataban(intI).Contains("R") Then
                                        intTRCnt = intTRCnt + Int(strInpUse(intI))
                                    Else
                                        intTLCnt = intTLCnt + Int(strInpUse(intI))
                                    End If

                                ElseIf strSelKataban(intI).Substring(5, 2) = "T6" Then
                                    intT6Cnt = intT6Cnt + Int(strInpUse(intI))

                                ElseIf strSelKataban(intI).Substring(5, 2) = "T7" Then
                                    'intT7Cnt = intT7Cnt + Int(strInpUse(intI)) 2016/08/19 RM1608024 K.Ohwaki Define
                                    If strSelKataban(intI).Substring(5, 4) = "T7EC" Then
                                        intT7ECCnt = intT7ECCnt + Int(strInpUse(intI))
                                    Else
                                        intT7Cnt = intT7Cnt + Int(strInpUse(intI))
                                    End If
                                    'End 2016/08/19 RM1608024 K.Ohwaki
                                End If
                            End If
                        Next

                        For intI As Integer = CdCst.Siyou_01.Valve1 - 1 To CdCst.Siyou_01.Valve7 - 1
                            intValveCnt = intValveCnt + Int(strInpUse(intI))
                            If Int(strInpUse(intI)) > 0 Then
                                Select Case strSelKataban(intI).PadRight(5, " ").Substring(0, 5)
                                    Case "N3E00", "N4E00"
                                        intValveCnt07 = intValveCnt07 + Int(strInpUse(intI))
                                    Case Else
                                        intValveCnt10 = intValveCnt10 + Int(strInpUse(intI))
                                End Select
                            End If
                        Next

                        For intI As Integer = CdCst.Siyou_01.Exhaust1 - 1 To CdCst.Siyou_01.Exhaust4 - 1
                            intExhaustCnt = intExhaustCnt + Int(strInpUse(intI))
                        Next

                        For intI As Integer = CdCst.Siyou_01.Regulat1 - 1 To CdCst.Siyou_01.Regulat2 - 1
                            intReguCnt = intReguCnt + Int(strInpUse(intI))
                        Next
                        intEndLCnt = intEndLCnt + Int(strInpUse(CdCst.Siyou_01.EndL - 1))
                        intEndRCnt = intEndRCnt + Int(strInpUse(CdCst.Siyou_01.EndR - 1))
                        'ﾀﾞﾐｰﾌﾞﾛｯｸ使用数をカウント
                        For intI As Integer = CdCst.Siyou_01.Dummy1 - 1 To CdCst.Siyou_01.Dummy2 - 1
                            intDummyCnt = intDummyCnt + Int(strInpUse(intI))
                        Next
                    End If

                    '計算
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "MN3Q0"
                            If objKtbnStrc.strcSelection.strOpSymbol(5) = "TX" Then
                                '左側＋右側電装ブロック(TX)の場合
                                objKtbnStrc.strcSelection.decDinRailLength = (intValveCnt10 * 10.5) + (intExhaustCnt * 12.5) + +64
                            Else
                                '左側もしくは右側電装ブロック(TX)の場合
                                objKtbnStrc.strcSelection.decDinRailLength = (intValveCnt10 * 10.5) + (intExhaustCnt * 12.5) + +53
                            End If
                        Case "MT3Q0"
                            'MT3Q0シリーズはマニホールド長さ(L1)不要のため、０をセットする
                            objKtbnStrc.strcSelection.decDinRailLength = 0
                        Case Else
                            objKtbnStrc.strcSelection.decDinRailLength = (intValveCnt07 * 7) + (intValveCnt10 * 10) + _
                                        (intExhaustCnt * 15.5) + (intEndLCnt * 15.3) + _
                                        (intEndRCnt * 15.9) + (intTMCnt * 12) + _
                                        (intTLCnt * 26.5) + (intTRCnt * 27.1) + _
                                        (intT6Cnt * 99.7) + (intT7Cnt * 57.2) + _
                                        (intReguCnt * 30) + (intDummyCnt * 7) + _
                                        (intT7ECCnt * 70.7) '2016/08/23 RM1608024 K.Ohwaki Append
                    End Select
                    '自動計算値保持
                    dblX = Int(((objKtbnStrc.strcSelection.decDinRailLength + 25) / 12.5) * 10)
                    dblX = Int(dblX * -0.1) * -1
                    dblStdNum = dblX * 12.5

                    '取付レール未変更時、または入力が0の場合、自動計算する
                    If strRailChangeFlg = "1" AndAlso intSelRail > 0 Then
                        '手動入力した時
                        dblRailLen = intSelRail
                    Else
                        '取付レール未変更時、または入力が0の場合、自動計算する
                        dblRailLen = dblStdNum
                    End If
                Case 2
                    If strInpUse Is Nothing Then
                    ElseIf strInpUse.Count = 0 Then
                    Else
                        '給排気ブロック使用数カウント
                        For intI As Integer = CdCst.Siyou_02.Exhaust1 - 1 To CdCst.Siyou_02.Exhaust6 - 1
                            If strInpUse(intI) > 0 Then
                                intExhaustCnt = intExhaustCnt + strInpUse(intI)
                            End If
                        Next
                        'バルブブロック使用数カウント
                        For intI As Integer = CdCst.Siyou_02.Valve1 - 1 To CdCst.Siyou_02.Valve6 - 1
                            If strInpUse(intI) > 0 Then
                                intValveCnt = intValveCnt + strInpUse(intI)
                            End If
                        Next
                        '仕切りブロック使用数カウント
                        For inti As Integer = CdCst.Siyou_02.Partition1 - 1 To CdCst.Siyou_02.Partition2 - 1
                            If strInpUse(inti) > 0 Then
                                intPositionCnt = intPositionCnt + strInpUse(inti)
                            End If
                        Next

                    End If

                    ''マニホールド長さ計算
                    If objKtbnStrc.strcSelection.strSeriesKataban = "MN4KB1" Then
                        dblStdNum = (intValveCnt * 16) + (intExhaustCnt * 16) + (intPositionCnt * 8) + 40
                    Else
                        dblStdNum = (intValveCnt * 19) + (intExhaustCnt * 20) + (intPositionCnt * 8) + 40
                    End If

                    'チェック用基数を取付レール長さ値に変更
                    dblX = Fix((dblStdNum + 40) / 12.5)
                    dblY = dblStdNum - dblX
                    If dblY = 0 Then
                        dblZ = dblX
                    Else
                        dblZ = dblX + 1
                    End If
                    dblRailLen = dblZ * 12.5
                    dblStdNum = dblRailLen

                    If strRailChangeFlg = "1" AndAlso intSelRail > 0 Then
                        '手動入力した時
                        dblRailLen = intSelRail
                    Else
                        '取付レール未変更時、または入力が0の場合、自動計算する
                        dblRailLen = dblStdNum
                    End If
                Case 3
                    dblRailLen = GetMonifold3_Data(objKtbnStrc)
                    dblStdNum = dblRailLen
                Case 4
                    Dim intWide As Double = GetMonifold4_Data(objKtbnStrc)
                    '計算
                    For inti As Integer = 0 To 79
                        If inti = 0 Then
                            If intWide >= inti * 12.5 And intWide <= (inti + 1) * 12.5 Then
                                dblRailLen = (inti + 1) * 12.5
                            End If
                            'strValues.Add((inti + 1) * 12.5)
                        Else
                            If intWide > inti * 12.5 And intWide <= (inti + 1) * 12.5 Then
                                dblRailLen = (inti + 1) * 12.5
                            End If
                            'strValues.Add((inti + 1) * 12.5)
                        End If
                    Next
                Case 7
                    Dim intMixCnt As Integer = 0
                    Dim dblManiLen As Double = 0D
                    Dim strDen As String = String.Empty
                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                    '   objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" then
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Or _
                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Or _
                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "V" Then 'RM1610013
                        strDen = objKtbnStrc.strcSelection.strOpSymbol(5)
                    Else
                        strDen = objKtbnStrc.strcSelection.strOpSymbol(4)
                    End If

                    If strInpUse Is Nothing Then

                    ElseIf strInpUse.Count = 0 Then
                    Else
                        'ミックスブロック使用数カウント
                        intMixCnt = CInt(strInpUse(CdCst.Siyou_07.Mix - 1))
                        '仕切ブロック使用数カウント
                        For intI As Integer = CdCst.Siyou_07.Partition1 - 1 To CdCst.Siyou_07.Partition2 - 1
                            intPartCnt = intPartCnt + CInt(strInpUse(intI))
                        Next
                        Select Case objKtbnStrc.strcSelection.strSeriesKataban
                            Case "MN3GA1", "MN4GA1", "MN3GB1", "MN4GB1"
                                'バルブブロック使用数カウント
                                For intI As Integer = CdCst.Siyou_07.Elect1 - 1 To CdCst.Siyou_07.Elect8 - 1
                                    intValveCnt1 = intValveCnt1 + CInt(strInpUse(intI))
                                Next
                                '給排気ブロック使用数カウント
                                For intI As Integer = CdCst.Siyou_07.Exhaust1 - 1 To CdCst.Siyou_07.Exhaust3 - 1
                                    intExhaustCnt1 = intExhaustCnt1 + CInt(strInpUse(intI))
                                Next
                                'マニホールド長さ
                                dblManiLen = (intValveCnt1 * 10.5) + (intExhaustCnt1 * 16) + (intPartCnt * 10.5)
                                If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                   objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    Select Case Left(strDen.ToString, 2)
                                        Case "T1"
                                            dblManiLen = dblManiLen + 83.8
                                        Case "T3", "T5"
                                            dblManiLen = dblManiLen + 69.3
                                        Case "T6"
                                            dblManiLen = dblManiLen + 143.5
                                        Case "T7", "T8"
                                            dblManiLen = dblManiLen + 64.3
                                        Case Else
                                            dblManiLen = dblManiLen + 41
                                    End Select
                                Else
                                    Select Case Left(strDen.ToString, 2)
                                        Case "T1"
                                            dblManiLen = dblManiLen + 87
                                        Case "T3", "T5"
                                            dblManiLen = dblManiLen + 72.5
                                        Case "T6"
                                            dblManiLen = dblManiLen + 144
                                        Case "T7", "T8"
                                            dblManiLen = dblManiLen + 67.5
                                        Case Else
                                            dblManiLen = dblManiLen + 42
                                    End Select
                                End If
                                objKtbnStrc.strcSelection.decDinRailLength = dblManiLen
                            Case "MN3GA2", "MN4GA2", "MN3GB2", "MN4GB2"
                                'バルブブロック使用数カウント
                                For intI As Integer = CdCst.Siyou_07.Elect1 - 1 To CdCst.Siyou_07.Elect8 - 1
                                    intValveCnt1 = intValveCnt1 + CInt(strInpUse(intI))
                                Next
                                '給排気ブロック使用数カウント
                                For intI As Integer = CdCst.Siyou_07.Exhaust1 - 1 To CdCst.Siyou_07.Exhaust3 - 1
                                    intExhaustCnt1 = intExhaustCnt1 + CInt(strInpUse(intI))
                                Next
                                'マニホールド長さ
                                dblManiLen = (intValveCnt1 * 16) + (intExhaustCnt1 * 18) + (intPartCnt * 10.5)
                                If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                   objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    Select Case Left(strDen.ToString, 2)
                                        Case "T1"
                                            dblManiLen = dblManiLen + 86.3
                                        Case "T3", "T5"
                                            dblManiLen = dblManiLen + 71.8
                                        Case "T6"
                                            dblManiLen = dblManiLen + 146
                                        Case "T7", "T8"
                                            dblManiLen = dblManiLen + 66.8
                                        Case Else
                                            dblManiLen = dblManiLen + 46
                                    End Select
                                Else
                                    Select Case Left(strDen.ToString, 2)
                                        Case "T1"
                                            dblManiLen = dblManiLen + 89.5
                                        Case "T3", "T5"
                                            dblManiLen = dblManiLen + 75
                                        Case "T6"
                                            dblManiLen = dblManiLen + 146.5
                                        Case "T7", "T8"
                                            dblManiLen = dblManiLen + 70
                                        Case Else
                                            dblManiLen = dblManiLen + 47
                                    End Select
                                End If
                                objKtbnStrc.strcSelection.decDinRailLength = dblManiLen
                            Case "MN3GAX12", "MN4GAX12", "MN3GBX12", "MN4GBX12", "MN4GDX12", "MN4GEX12"             'RM1303003 2013/03/05
                                'バルブブロック使用数カウント
                                For intI As Integer = CdCst.Siyou_07.Elect1 - 1 To CdCst.Siyou_07.Elect8 - 1
                                    If Left(strSelKataban(intI) & Space(5), 5).Substring(4, 1) = "1" Then
                                        intValveCnt1 = intValveCnt1 + CInt(strInpUse(intI))
                                    ElseIf Left(strSelKataban(intI) & Space(5), 5).Substring(4, 1) = "2" Then
                                        intValveCnt2 = intValveCnt2 + CInt(strInpUse(intI))
                                    End If
                                Next
                                '給排気ブロック使用数カウント
                                For intI As Integer = CdCst.Siyou_07.Exhaust1 - 1 To CdCst.Siyou_07.Exhaust3 - 1
                                    If Left(strSelKataban(intI) & Space(4), 4).Substring(3, 1) = "1" Then
                                        intExhaustCnt1 = intExhaustCnt1 + CInt(strInpUse(intI))
                                    ElseIf Left(strSelKataban(intI) & Space(4), 4).Substring(3, 1) = "2" Then
                                        intExhaustCnt2 = intExhaustCnt2 + CInt(strInpUse(intI))
                                    End If
                                Next
                                'マニホールド長さ
                                dblManiLen = (intValveCnt1 * 10.5) + (intValveCnt2 * 16) + _
                                             (intExhaustCnt1 * 16) + (intExhaustCnt2 * 18) + _
                                             (intPartCnt * 10.5) + (intMixCnt * 16)
                                If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                                   objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                    Select Case Left(strDen.ToString, 2)
                                        Case "T1"
                                            If InStr(strDen.ToString, "R") > 0 Then
                                                dblManiLen = dblManiLen + 83.8
                                            Else
                                                dblManiLen = dblManiLen + 86.3
                                            End If
                                        Case "T3", "T5"
                                            If InStr(strDen.ToString, "R") > 0 Then
                                                dblManiLen = dblManiLen + 69.3
                                            Else
                                                dblManiLen = dblManiLen + 71.8
                                            End If
                                        Case "T6"
                                            dblManiLen = dblManiLen + 146
                                        Case "T7", "T8"
                                            dblManiLen = dblManiLen + 66.8
                                        Case Else
                                            dblManiLen = dblManiLen + 44.5
                                    End Select
                                Else
                                    Select Case Left(strDen.ToString, 2)
                                        Case "T1"
                                            If InStr(strDen.ToString, "R") > 0 Then
                                                dblManiLen = dblManiLen + 87
                                            Else
                                                dblManiLen = dblManiLen + 89.5
                                            End If
                                        Case "T3", "T5"
                                            If InStr(strDen.ToString, "R") > 0 Then
                                                dblManiLen = dblManiLen + 72.5
                                            Else
                                                dblManiLen = dblManiLen + 75
                                            End If
                                        Case "T6"
                                            dblManiLen = dblManiLen + 146.5
                                        Case "T7", "T8"
                                            dblManiLen = dblManiLen + 70
                                        Case Else
                                            dblManiLen = dblManiLen + 44.5
                                    End Select
                                End If
                                objKtbnStrc.strcSelection.decDinRailLength = dblManiLen
                        End Select
                    End If

                    dblStdNum = Fix(((dblManiLen + 40) / 12.5) + 0.99) * 12.5

                    '取付レール未変更時、または入力が0の場合、自動計算する
                    If strRailChangeFlg = "1" AndAlso intSelRail > 0 Then
                        '手動入力した時
                        dblRailLen = intSelRail
                    Else
                        '取付レール未変更時、または入力が0の場合、自動計算する
                        dblRailLen = dblStdNum
                    End If
                Case 8
                    If strInpUse Is Nothing Then
                    ElseIf strInpUse.Count = 0 Then
                    Else
                        '使用数カウント
                        For intI As Integer = CdCst.Siyou_08.ElType1 - 1 To CdCst.Siyou_08.Exhaust2 - 1
                            If strInpUse(intI) > 0 Then
                                intExhaustCnt = intExhaustCnt + strInpUse(intI)
                            End If
                        Next
                    End If

                    intExhaustCnt = intExhaustCnt + 2

                    ''取付レール長さ計算
                    '取付レール未変更時、または入力が0の場合、自動計算する
                    If strRailChangeFlg = "1" AndAlso intSelRail > 0 Then
                        '手動入力した時
                        dblRailLen = intSelRail
                        dblStdNum = intExhaustCnt * 16 + 32
                    Else
                        '取付レール未変更時、または入力が0の場合、自動計算する
                        dblRailLen = intExhaustCnt * 16 + 32
                        dblStdNum = dblRailLen
                    End If
                Case 10
                    Dim intValve As Integer
                    Dim intExhlt As Integer
                    Dim intPart As Integer
                    Dim intLen As Integer
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "MN3S0", "MN4S0"
                            '使用数ｶｳﾝﾄ
                            For idx As Integer = 1 To 11
                                Select Case idx
                                    Case 1, 2, 3, 4, 5, 6, 7
                                        intValve = intValve + CInt(strInpUse(idx))
                                    Case 8, 9
                                        intExhlt = intExhlt + CInt(strInpUse(idx))
                                    Case 10, 11
                                        intPart = intPart + CInt(strInpUse(idx))
                                End Select
                            Next

                            '取付ﾚｰﾙ長さ計算
                            Dim dblManiLen As Double = (intValve * 11) + (intExhlt * 16) + (intPart * 6)

                            Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(6).ToString, 2)
                                Case "T1"
                                    dblManiLen += 107
                                Case "T3", "T5"
                                    dblManiLen += 57
                                Case "T6"
                                    dblManiLen += 128.5
                                Case Else
                                    dblManiLen += 42
                            End Select

                            dblManiLen = (dblManiLen + 40) / 12.5

                            intLen = CStr(dblManiLen).IndexOf(".")
                            If Not intLen = -1 Then
                                dblManiLen = CDec(Left(CStr(dblManiLen), intLen))
                                dblManiLen = dblManiLen + 1
                            End If

                            dblManiLen = dblManiLen * 12.5

                            If strRailChangeFlg = "1" Then dblManiLen = intSelRail

                            objKtbnStrc.strcSelection.decDinRailLength = dblManiLen
                            '画面表示値
                            dblRailLen = dblManiLen
                            dblStdNum = dblManiLen
                    End Select
                Case 11
                    Dim intTotalCharge As Integer

                    '集中給気ブロック／APS付集中給気ブロックの合計数
                    intTotalCharge = CInt(strInpUse(CdCst.Siyou_11.ChargeAir - 1)) + CInt(strInpUse(CdCst.Siyou_11.ChargeAirAPS - 1))
                    '取付レール長さ標準値取得
                    If objKtbnStrc.strcSelection.strOpSymbol(3) = "D" Then       '選択オプション(3)が"D"
                        dblStdNum = 0
                    Else                                            '選択オプション(3)が"D"以外
                        Select Case objKtbnStrc.strcSelection.strSeriesKataban
                            Case "MNRB500A", "MNRJB500A"
                                Select Case intTotalCharge
                                    Case 0
                                        dblStdNum = 0
                                    Case 1, 2
                                        dblStdNum = 25 * CInt(objKtbnStrc.strcSelection.strOpSymbol(2) - 1) + _
                                                    Int((CInt(objKtbnStrc.strcSelection.strOpSymbol(2)) + 1 - intTotalCharge) / 4) * 12.5 + _
                                                    100 + 25 * intTotalCharge
                                    Case 3
                                        dblStdNum = 25 * CInt(objKtbnStrc.strcSelection.strOpSymbol(2) - 1) + _
                                                    Int((CInt(objKtbnStrc.strcSelection.strOpSymbol(2)) - 2 + intTotalCharge) / 4) * 12.5 + _
                                                    125 + 12.5 * intTotalCharge
                                End Select
                            Case "MNRB500B", "MNRJB500B"
                                dblStdNum = 25 * CInt(objKtbnStrc.strcSelection.strOpSymbol(2) - 1) + Int((CInt(objKtbnStrc.strcSelection.strOpSymbol(2)) + 2) / 4) * 12.5 + 100
                        End Select
                    End If
                    '入力値がない場合、標準値使用
                    If strRailChangeFlg = "1" AndAlso intSelRail > 0 Then
                        '手動入力した時
                        dblRailLen = intSelRail
                    Else
                        '取付レール未変更時、または入力が0の場合、自動計算する
                        dblRailLen = dblStdNum
                    End If
                Case 13
                    Dim dblManiStd As Double
                    Dim strOpSymbol As String

                    If strInpUse Is Nothing Then
                    ElseIf strInpUse.Count = 0 Then
                    Else
                        'バルブブロック使用数カウント
                        For intI As Integer = CdCst.Siyou_13.Valve1 - 1 To CdCst.Siyou_13.Valve6 - 1
                            If CInt(strInpUse(intI)) > 0 Then
                                If strSelKataban(intI).Length > 0 Then
                                    If strSelKataban(intI).Substring(4, 1) = "1" Then
                                        intValveCnt1 = intValveCnt1 + CInt(strInpUse(intI))
                                    Else
                                        intValveCnt2 = intValveCnt2 + CInt(strInpUse(intI))
                                    End If
                                End If
                            End If
                        Next

                        '給排気ブロック使用数カウント
                        For intI As Integer = CdCst.Siyou_13.Exhaust1 - 1 To CdCst.Siyou_13.Exhaust6 - 1
                            If CInt(strInpUse(intI)) > 0 Then
                                intExhaustCnt = intExhaustCnt + CInt(strInpUse(intI))
                            End If
                        Next
                    End If

                    'マニホールド長さ計算
                    Dim dblManiLen As Double = 0D
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "MN4TB1"
                            dblManiLen = (intValveCnt1 * 17) + (intExhaustCnt * 17) + 40
                            strOpSymbol = objKtbnStrc.strcSelection.strOpSymbol(6).ToString
                        Case "MN4TB2"
                            dblManiLen = (intValveCnt2 * 20) + (intExhaustCnt * 20) + 40
                            strOpSymbol = objKtbnStrc.strcSelection.strOpSymbol(6).ToString
                        Case Else
                            dblManiLen = (intValveCnt1 * 17) + (intValveCnt2 * 20) + (intExhaustCnt * 20) + 40
                            strOpSymbol = objKtbnStrc.strcSelection.strOpSymbol(4).ToString
                    End Select

                    Select Case strOpSymbol
                        Case "T10", "T30", "T31", "T50"
                            dblManiLen += 20
                        Case Else
                            dblManiLen += 60
                    End Select
                    objKtbnStrc.strcSelection.decDinRailLength = dblManiLen

                    '取付レール長さ基数
                    dblManiStd = Fix((dblManiLen - 110 + 12.499) / 12.5)

                    '取付レール長さ計算
                    dblStdNum = dblManiStd * 12.5 + 150
                    '取付レール未変更時、または入力が0の場合、自動計算する
                    If strRailChangeFlg = "1" AndAlso intSelRail > 0 Then
                        '手動入力した時
                        dblRailLen = intSelRail
                    Else
                        '取付レール未変更時、または入力が0の場合、自動計算する
                        dblRailLen = dblStdNum
                    End If
                Case 14
                    Dim intEvtCnt As Integer = 0
                    Dim intEndCnt As Integer = 0
                    Dim intX As Integer

                    '10.1 使用数カウント
                    'EVT
                    intEvtCnt = strInpUse(CdCst.Siyou_14.Evt - 1)
                    '電装・給気ブロック
                    For intI As Integer = CdCst.Siyou_14.Exhaust1 To CdCst.Siyou_14.Exhaust3
                        intExhaustCnt = intExhaustCnt + CInt(strInpUse(intI - 1))
                    Next
                    'エンドブロック
                    For intI As Integer = CdCst.Siyou_14.End1 To CdCst.Siyou_14.End2
                        intEndCnt = intEndCnt + CInt(strInpUse(intI - 1))
                    Next

                    '10.2 マニホールド長さ計算
                    Dim dblManiLen As Double = 0D
                    Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(4), 2)
                        Case "T1", "T3"
                            dblManiLen = (intEvtCnt * 14) + (intEndCnt * 10) + 23 + (intExhaustCnt * 42)
                        Case "T9"
                            dblManiLen = (intEvtCnt * 14) + (intEndCnt * 10) + 23 + (intExhaustCnt * 32)
                    End Select
                    objKtbnStrc.strcSelection.decDinRailLength = dblManiLen
                    '10.3 取付レール長さ計算
                    dblX = (dblManiLen + 40) / 12.5
                    intX = Fix(dblX)
                    dblY = intX - dblX
                    If dblY = 0 Then
                    Else
                        intX = intX + 1
                    End If
                    dblStdNum = intX * 12.5

                    '取付レール未変更時、または入力が0の場合、自動計算値適用
                    If strRailChangeFlg = "1" AndAlso intSelRail > 0 Then
                        '手動入力した時
                        dblRailLen = intSelRail
                    Else
                        '取付レール未変更時、または入力が0の場合、自動計算する
                        dblRailLen = dblStdNum
                    End If
                Case 15
                    Dim intInOutCnt As Integer = 0
                    Dim intExhauCnt As Integer = 0
                    Dim dblWiring As Double = 0
                    Dim sbValues As New System.Text.StringBuilder

                    If objKtbnStrc.strcSelection.strOpSymbol(7).ToString = "D" Then
                        For intI As Integer = CdCst.Siyou_15.InOut1 To CdCst.Siyou_15.InOut2
                            intInOutCnt = intInOutCnt + Int(strInpUse(intI - 1))
                        Next
                        For intI As Integer = CdCst.Siyou_15.Valve1 To CdCst.Siyou_15.Valve8
                            intValveCnt = intValveCnt + Int(strInpUse(intI - 1))
                        Next
                        For intI As Integer = CdCst.Siyou_15.Exhaust1 To CdCst.Siyou_15.Exhaust2
                            intExhauCnt = intExhauCnt + Int(strInpUse(intI - 1))
                        Next
                        For intI As Integer = CdCst.Siyou_15.Partition1 To CdCst.Siyou_15.Partition2
                            intPartCnt = intPartCnt + Int(strInpUse(intI - 1))
                        Next

                        Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(4).ToString, 2)
                            Case "T1"
                                dblWiring = 175.5
                                dblR1 = 0
                            Case "T2"
                                dblWiring = 110
                                dblR1 = 0
                            Case "T3", "T5"
                                dblWiring = 106
                                dblR1 = 0
                            Case "T8"
                                dblWiring = 148.5
                                dblR1 = 0
                            Case "R1"
                                If objKtbnStrc.strcSelection.strFullKataban.Contains("MW4GA") Or _
                                    objKtbnStrc.strcSelection.strFullKataban.Contains("MW3GA") Or _
                                    objKtbnStrc.strcSelection.strFullKataban.Contains("MW4GB") Then
                                    dblWiring = 51
                                    dblR1 = 24
                                Else
                                    dblWiring = 51
                                    dblR1 = 0
                                End If
                        End Select

                        'マニホールド長さ基数計算
                        If dblWiring > 0 Then
                            dblStdNum = (intValveCnt * 16) + (intExhauCnt * 18) + (intPartCnt * 13.5) + _
                                        (intInOutCnt * 45) + dblWiring
                        End If

                        '取付レール長さ計算
                        If strRailChangeFlg = "1" AndAlso intSelRail > 0 Then
                            '手動入力した時
                            dblRailLen = intSelRail
                        Else
                            '取付レール未変更時、または入力が0の場合、自動計算する
                            dblX = ((dblStdNum + 40) / 12.5) * 10 + dblR1
                            dblX = Math.Ceiling(Int(dblX * -0.1) * -1)
                            dblRailLen = dblX * 12.5
                        End If
                    Else
                        dblRailLen = 0
                    End If
                Case 18
                    Dim strDen As String = String.Empty
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                       objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                        strDen = objKtbnStrc.strcSelection.strOpSymbol(5)
                    Else
                        strDen = objKtbnStrc.strcSelection.strOpSymbol(4)
                    End If
                    '仕切ブロック使用数カウント
                    For intI As Integer = CdCst.Siyou_18.Partition1 - 1 To CdCst.Siyou_18.Partition2 - 1
                        intPartCnt = intPartCnt + CInt(strInpUse(intI))
                    Next

                    Dim dblManiLen As Double = 0D
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "MN3GD1", "MN4GD1", "MN3GE1", "MN4GE1"
                            'バルブブロック使用数カウント
                            For intI As Integer = CdCst.Siyou_18.Elect1 - 1 To CdCst.Siyou_18.Elect8 - 1
                                intValveCnt1 = intValveCnt1 + CInt(strInpUse(intI))
                            Next
                            '給排気ブロック使用数カウント
                            For intI As Integer = CdCst.Siyou_18.Exhaust1 - 1 To CdCst.Siyou_18.Exhaust3 - 1
                                intExhaustCnt1 = intExhaustCnt1 + CInt(strInpUse(intI))
                            Next
                            'マニホールド長さ
                            dblManiLen = (intValveCnt1 * 10.5) + (intExhaustCnt1 * 16) + (intPartCnt * 10.5)
                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                               objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                Select Case Strings.Left(strDen, 2)
                                    Case "T1"
                                        dblManiLen = dblManiLen + 83.8
                                    Case "T3", "T5"
                                        dblManiLen = dblManiLen + 69.3
                                    Case "T6"
                                        dblManiLen = dblManiLen + 143.5
                                    Case "T7", "T8"
                                        dblManiLen = dblManiLen + 64.3
                                    Case Else
                                        dblManiLen = dblManiLen + 41
                                End Select
                            Else
                                Select Case Strings.Left(strDen, 2)
                                    Case "T1"
                                        dblManiLen = dblManiLen + 87
                                    Case "T3", "T5"
                                        dblManiLen = dblManiLen + 72.5
                                    Case "T6"
                                        dblManiLen = dblManiLen + 144
                                    Case "T7", "T8"
                                        dblManiLen = dblManiLen + 67.5
                                    Case Else
                                        dblManiLen = dblManiLen + 42
                                End Select
                            End If

                        Case "MN3GD2", "MN4GD2", "MN3GE2", "MN4GE2"
                            'バルブブロック使用数カウント
                            For intI As Integer = CdCst.Siyou_18.Elect1 - 1 To CdCst.Siyou_18.Elect8 - 1
                                intValveCnt1 = intValveCnt1 + CInt(strInpUse(intI))
                            Next
                            '給排気ブロック使用数カウント
                            For intI As Integer = CdCst.Siyou_18.Exhaust1 - 1 To CdCst.Siyou_18.Exhaust3 - 1
                                intExhaustCnt1 = intExhaustCnt1 + CInt(strInpUse(intI))
                            Next
                            'マニホールド長さ
                            dblManiLen = (intValveCnt1 * 16) + (intExhaustCnt1 * 18) + (intPartCnt * 10.5)
                            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "R" Or _
                               objKtbnStrc.strcSelection.strKeyKataban.Trim = "U" Then
                                Select Case Strings.Left(strDen, 2)
                                    Case "T1"
                                        dblManiLen = dblManiLen + 86.3
                                    Case "T3", "T5"
                                        dblManiLen = dblManiLen + 71.8
                                    Case "T6"
                                        dblManiLen = dblManiLen + 146
                                    Case "T7", "T8"
                                        dblManiLen = dblManiLen + 66.8
                                    Case Else
                                        dblManiLen = dblManiLen + 46
                                End Select
                            Else
                                Select Case Strings.Left(strDen, 2)
                                    Case "T1"
                                        dblManiLen = dblManiLen + 89.5
                                    Case "T3", "T5"
                                        dblManiLen = dblManiLen + 75
                                    Case "T6"
                                        dblManiLen = dblManiLen + 146.5
                                    Case "T7", "T8"
                                        dblManiLen = dblManiLen + 70
                                    Case Else
                                        dblManiLen = dblManiLen + 47
                                End Select
                            End If
                    End Select
                    objKtbnStrc.strcSelection.decDinRailLength = dblManiLen
                    '取付レール未変更時、または入力が0の場合、自動計算する
                    dblStdNum = Fix(((dblManiLen + 40) / 12.5) + 0.99) * 12.5
                    If strRailChangeFlg = "1" AndAlso intSelRail > 0 Then
                        '手動入力した時
                        dblRailLen = intSelRail
                    Else
                        '取付レール未変更時、または入力が0の場合、自動計算する
                        dblRailLen = dblStdNum
                    End If
            End Select
        Catch ex As Exception
            Call WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 取付レール値計算・取得
    ''' </summary>
    ''' <param name="ManifoldMode"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function subGetRail_Cmb(ByVal ManifoldMode As Integer, ByRef objKtbnStrc As KHKtbnStrc) As ArrayList
        Dim dblX As Double
        subGetRail_Cmb = New ArrayList
        Try
            Select Case ManifoldMode
                Case 1
                    If objKtbnStrc.strcSelection.strSeriesKataban <> "MT3Q0" Then 'MT3Q0シリーズは取付レール長さなし
                        '取付レール長さドロップダウン
                        dblX = 25
                        For intI As Integer = 1 To 79
                            subGetRail_Cmb.Add(dblX)
                            dblX = dblX + 12.5
                        Next
                    End If
                Case 2
                    dblX = 25
                    For intI As Integer = 1 To 79
                        subGetRail_Cmb.Add(CStr(dblX))
                        dblX = dblX + 12.5
                    Next
                Case 3
                    Dim dblRailLen As Double = GetMonifold3_Data(objKtbnStrc)
                    If dblRailLen > 0 Then subGetRail_Cmb.Add(dblRailLen)
                Case 4
                    Dim intWide As Double = GetMonifold4_Data(objKtbnStrc)
                    '２０行目変更不可
                    Dim strUL As String = String.Empty
                    Select Case Strings.Right(objKtbnStrc.strcSelection.strSeriesKataban, 1).ToString
                        Case "1", "2"
                            strUL = objKtbnStrc.strcSelection.strOpSymbol(12).ToString
                        Case "3", "4"
                            strUL = objKtbnStrc.strcSelection.strOpSymbol(11).ToString
                    End Select
                    'If strUL = "UL" Then intWide = 0
                    '計算
                    If intWide > 0 Then
                        '2016/10/05　RM1609066　12.5～500 12.5刻み　→　87.5～1000 12.5刻み
                        'For inti As Integer = 0 To 39
                        For inti As Integer = 6 To 79
                            If inti = 0 Then
                                If intWide >= inti * 12.5 And intWide <= (inti + 1) * 12.5 Then
                                    dblX = (inti + 1) * 12.5
                                End If
                                subGetRail_Cmb.Add((inti + 1) * 12.5)
                            Else
                                If intWide > inti * 12.5 And intWide <= (inti + 1) * 12.5 Then
                                    dblX = (inti + 1) * 12.5
                                End If
                                subGetRail_Cmb.Add((inti + 1) * 12.5)
                            End If
                        Next
                    End If
                Case 7
                    '取付レール長さドロップダウン
                    'dblX = 25
                    'For intI As Integer = 1 To 50
                    '2016/10/05　RM1609066　25～637.5 12.5刻み　→　87.5～1000 12.5刻み
                    dblX = 87.5
                    For inti As Integer = 1 To 74
                        subGetRail_Cmb.Add(dblX)
                        dblX = dblX + 12.5
                    Next
                Case 8
                    dblX = 64
                    For intI As Integer = 1 To 50
                        subGetRail_Cmb.Add(CStr(dblX))
                        dblX = dblX + 16
                    Next
                Case 10
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "MN3S0", "MN4S0"
                            dblX = 87.5
                            For idx As Integer = 1 To 30
                                subGetRail_Cmb.Add(dblX)
                                dblX = dblX + 12.5
                            Next
                    End Select
                Case 11
                    If objKtbnStrc.strcSelection.strOpSymbol(3) <> "D" Then       '選択オプション(3)が"D"
                        dblX = 0
                        For idx As Integer = 1 To 48
                            dblX += 12.5
                            subGetRail_Cmb.Add(dblX)
                        Next
                    End If
                Case 13
                    dblX = 112.5
                    For intI As Integer = 1 To 30
                        subGetRail_Cmb.Add(dblX)
                        dblX = dblX + 12.5
                    Next
                Case 14
                    dblX = 75
                    For intI As Integer = 1 To 50
                        subGetRail_Cmb.Add(dblX)
                        dblX = dblX + 12.5
                    Next
                Case 15
                    If objKtbnStrc.strcSelection.strOpSymbol(7).ToString = "D" Then
                        '取付レール長さドロップダウン
                        dblX = 25
                        For intI As Integer = 1 To 79
                            subGetRail_Cmb.Add(dblX)
                            dblX = dblX + 12.5
                        Next
                    End If
                Case 18
                    '取付レール長さドロップダウン
                    dblX = 25
                    For intI As Integer = 1 To 79
                        subGetRail_Cmb.Add(dblX)
                        dblX = dblX + 12.5
                    Next
            End Select
        Catch ex As Exception
            Call WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' マニホールド3の場合
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetMonifold3_Data(objKtbnStrc As KHKtbnStrc) As Double
        GetMonifold3_Data = 0D
        Try
            If objKtbnStrc.strcSelection.strOpSymbol(1) = "D" Then    'オプション記号リスト(1)が"D"
                '連数値
                Dim intElectSeq As Integer = CInt(objKtbnStrc.strcSelection.strOpSymbol(10))
                If Left(objKtbnStrc.strcSelection.strOpSymbol(9), 1) = "T" Then   'オプション記号リスト(9)左１文字が"T"
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(9), 2) = "T3" Or _
                        Left(objKtbnStrc.strcSelection.strOpSymbol(9), 2) = "T5" Then
                        '取付レール長さ
                        GetMonifold3_Data = 125 + (intElectSeq - Int((intElectSeq + 1) / 6)) * 12.5
                    Else
                        '取付レール長さ
                        GetMonifold3_Data = 187.5 + (intElectSeq - Int((intElectSeq - 1) / 6)) * 12.5
                    End If
                Else                                                'オプション記号リスト(9)左１文字が"T"以外
                    '取付レール長さ
                    GetMonifold3_Data = 100 + (intElectSeq - Int((intElectSeq - 1) / 6)) * 12.5
                End If
            End If
        Catch ex As Exception
            Call WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' マニホールド4の場合
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetMonifold4_Data(objKtbnStrc As KHKtbnStrc) As Integer
        GetMonifold4_Data = 0
        Try
            Dim GetMani4Data As Hashtable = GetOptionData(objKtbnStrc, 4)
            Dim strElecConType As String = GetMani4Data("strElecConType")
            Dim strMaxSeq As String = GetMani4Data("strMaxSeq")
            Dim strStdMFType As String = GetMani4Data("strStdMFType")
            Dim strOptionD As String = GetMani4Data("strOptionD")
            Dim strPortSize As String = GetMani4Data("strPortSize")

            If strOptionD = "D" Then
                Select Case objKtbnStrc.strcSelection.strSeriesKataban
                    Case "M3GA1", "M4GA1", "M3GD1", "M4GD1"
                        If Left(strElecConType.Trim, 1) <> "T" Or _
                           Len(strElecConType.Trim) = 0 Then
                            '4GA/個別配線
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (10.5 * Val(Trim(strMaxSeq))) + 24.3 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (10.5 * Val(Trim(strMaxSeq))) + 29.3 + 40
                            End If
                        ElseIf Left(strElecConType.Trim, 2) = "T1" Or _
                               Left(strElecConType.Trim, 2) = "T3" Or _
                               Left(strElecConType.Trim, 2) = "T5" Then
                            '4GA/省配線(T1/T3/T5)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (12.5 * Val(Trim(strMaxSeq))) + 65.6 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (12.5 * Val(Trim(strMaxSeq))) + 70.6 + 40
                            End If
                        ElseIf Left(strElecConType.Trim, 2) = "T6" Then
                            '4GA/省配線(T6)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (12.5 * Val(Trim(strMaxSeq))) + 131.1 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (12.5 * Val(Trim(strMaxSeq))) + 136.1 + 40
                            End If
                            'RM1611037 T8分岐
                        ElseIf Left(strElecConType.Trim, 2) = "T8" Then
                            '4GA/省配線(T8)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (12.5 * Val(Trim(strMaxSeq))) + 67.1 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (12.5 * Val(Trim(strMaxSeq))) + 67.1 + 40
                            End If
                            'RM1611037 End
                        End If
                    Case "M3GA2", "M4GA2", "M3GD2", "M4GD2"
                        If Left(strElecConType.Trim, 1) <> "T" Or _
                           Len(strElecConType.Trim) = 0 Then
                            '4GA/個別配線
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (16 * Val(Trim(strMaxSeq))) + 22 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (16 * Val(Trim(strMaxSeq))) + 25 + 40
                            End If
                        ElseIf Left(strElecConType.Trim, 2) = "T1" Or _
                               Left(strElecConType.Trim, 2) = "T3" Or _
                               Left(strElecConType.Trim, 2) = "T5" Then
                            '4GA/省配線(T1/T3/T5)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (17.5 * Val(Trim(strMaxSeq))) + 64 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (17.5 * Val(Trim(strMaxSeq))) + 64 + 40
                            End If
                        ElseIf Left(strElecConType.Trim, 2) = "T6" Then
                            '4GA/省配線(T6)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (17.5 * Val(Trim(strMaxSeq))) + 129.5 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (17.5 * Val(Trim(strMaxSeq))) + 129.5 + 40
                            End If
                            'RM1611037 T8分岐
                        ElseIf Left(strElecConType.Trim, 2) = "T8" Then
                            '4GA/省配線(T8)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (17.5 * Val(Trim(strMaxSeq))) + 60.8 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (17.5 * Val(Trim(strMaxSeq))) + 60.8 + 40
                            End If
                            'RM1611037 End
                        End If
                    Case "M3GA3", "M4GA3", "M3GD3", "M4GD3"
                        If Left(strElecConType.Trim, 1) <> "T" Or Len(strElecConType.Trim) = 0 Then
                            '4GA/個別配線
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (19 * Val(Trim(strMaxSeq))) + 24 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (19 * Val(Trim(strMaxSeq))) + 28 + 40
                            End If
                        ElseIf Left(strElecConType.Trim, 2) = "T1" Or _
                               Left(strElecConType.Trim, 2) = "T3" Or _
                               Left(strElecConType.Trim, 2) = "T5" Then
                            '4GA/省配線(T1/T3/T5)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (20.5 * Val(Trim(strMaxSeq))) + 64.9 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (20.5 * Val(Trim(strMaxSeq))) + 65.9 + 40
                            End If
                        ElseIf Left(strElecConType.Trim, 2) = "T6" Then
                            '4GA/省配線(T6)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (20.5 * Val(Trim(strMaxSeq))) + 130.4 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (20.5 * Val(Trim(strMaxSeq))) + 131.4 + 40
                            End If
                            'RM1611037 T8分岐
                        ElseIf Left(strElecConType.Trim, 2) = "T8" Then
                            '4GA/省配線(T8)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (20.5 * Val(Trim(strMaxSeq))) + 62.7 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (20.5 * Val(Trim(strMaxSeq))) + 62.7 + 40
                            End If
                            'RM1611037 End
                        End If
                    Case "M3GB1", "M4GB1", "M3GE1", "M4GE1"
                        If Left(strElecConType.Trim, 2) = "T1" Or _
                               Left(strElecConType.Trim, 2) = "T3" Or _
                               Left(strElecConType.Trim, 2) = "T5" Then
                            '4GA/省配線(T1/T3/T5)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (12.5 * Val(Trim(strMaxSeq))) + 70.6 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (12.5 * Val(Trim(strMaxSeq))) + 70.6 + 40
                            End If
                        ElseIf Left(strElecConType.Trim, 2) = "T6" Then
                            '4GA/省配線(T6)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (12.5 * Val(Trim(strMaxSeq))) + 136.1 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (12.5 * Val(Trim(strMaxSeq))) + 136.1 + 40
                            End If
                            'RM1611037 T8分岐
                        ElseIf Left(strElecConType.Trim, 2) = "T8" Then
                            '4GA/省配線(T8)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (12.5 * Val(Trim(strMaxSeq))) + 67.1 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (12.5 * Val(Trim(strMaxSeq))) + 67.1 + 40
                            End If
                            'RM1611037 End
                        ElseIf Left(strElecConType.Trim, 1) <> "T" Or Len(strElecConType.Trim) = 0 Then
                            '4GA/個別配線
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                If strPortSize.Trim <> "C8" Then
                                    GetMonifold4_Data = (10.5 * Val(Trim(strMaxSeq))) + 29.3 + 40
                                Else
                                    GetMonifold4_Data = (12.5 * Val(Trim(strMaxSeq))) + 29.3 + 40
                                End If
                            Else
                                '外部パイロットＭＦ
                                If strPortSize.Trim <> "C8" Then
                                    GetMonifold4_Data = (10.5 * Val(Trim(strMaxSeq))) + 29.3 + 40
                                Else
                                    GetMonifold4_Data = (12.5 * Val(Trim(strMaxSeq))) + 29.3 + 40
                                End If
                            End If
                        End If
                    Case "M3GB2", "M4GB2", "M3GE2", "M4GE2"
                        If Left(strElecConType.Trim, 1) <> "T" Or Len(strElecConType.Trim) = 0 Then
                            '4GA/個別配線
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                If strPortSize.Trim <> "C10" Then
                                    GetMonifold4_Data = (16 * Val(Trim(strMaxSeq))) + 22 + 40
                                Else
                                    GetMonifold4_Data = (17.5 * Val(Trim(strMaxSeq))) + 22 + 40
                                End If
                            Else
                                '外部パイロットＭＦ
                                If strPortSize.Trim <> "C10" Then
                                    GetMonifold4_Data = (16 * Val(Trim(strMaxSeq))) + 25 + 40
                                Else
                                    GetMonifold4_Data = (17.5 * Val(Trim(strMaxSeq))) + 25 + 40
                                End If
                            End If
                        ElseIf Left(strElecConType.Trim, 2) = "T1" Or _
                               Left(strElecConType.Trim, 2) = "T3" Or _
                               Left(strElecConType.Trim, 2) = "T5" Then
                            '4GA/省配線(T1/T3/T5)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (17.5 * Val(Trim(strMaxSeq))) + 64 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (17.5 * Val(Trim(strMaxSeq))) + 64 + 40
                            End If
                        ElseIf Left(strElecConType.Trim, 2) = "T6" Then
                            '4GA/省配線(T6)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (17.5 * Val(Trim(strMaxSeq))) + 129.5 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (17.5 * Val(Trim(strMaxSeq))) + 129.5 + 40
                            End If
                            'RM1611037 T8分岐
                        ElseIf Left(strElecConType.Trim, 2) = "T8" Then
                            '4GA/省配線(T8)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (17.5 * Val(Trim(strMaxSeq))) + 60.8 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (17.5 * Val(Trim(strMaxSeq))) + 60.8 + 40
                            End If
                            'RM1611037 End
                        End If
                    Case "M3GB3", "M4GB3", "M3GE3", "M4GE3"
                        If Left(strElecConType.Trim, 1) <> "T" Or Len(strElecConType.Trim) = 0 Then
                            '4GA/個別配線
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (19 * Val(Trim(strMaxSeq))) + 26 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (19 * Val(Trim(strMaxSeq))) + 28 + 40
                            End If
                        ElseIf Left(strElecConType.Trim, 2) = "T1" Or _
                               Left(strElecConType.Trim, 2) = "T3" Or _
                               Left(strElecConType.Trim, 2) = "T5" Then
                            '4GA/省配線(T1/T3/T5)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (20.5 * Val(Trim(strMaxSeq))) + 65.9 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (20.5 * Val(Trim(strMaxSeq))) + 65.9 + 40
                            End If
                        ElseIf Left(strElecConType.Trim, 2) = "T6" Then
                            '4GA/省配線(T6)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (20.5 * Val(Trim(strMaxSeq))) + 131.4 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (20.5 * Val(Trim(strMaxSeq))) + 131.4 + 40
                            End If
                            'RM1611037 T8分岐
                        ElseIf Left(strElecConType.Trim, 2) = "T8" Then
                            '4GA/省配線(T8)
                            If strStdMFType.Trim <> "P" Then
                                '標準ＭＦ
                                GetMonifold4_Data = (20.5 * Val(Trim(strMaxSeq))) + 62.7 + 40
                            Else
                                '外部パイロットＭＦ
                                GetMonifold4_Data = (20.5 * Val(Trim(strMaxSeq))) + 62.7 + 40
                            End If
                            'RM1611037 End
                        End If
                    Case "M4GB4"
                        GetMonifold4_Data = (25 * Val(Trim(strMaxSeq))) + 99 + 40
                End Select
            End If
        Catch ex As Exception
            Call WriteErrorLog("E001", ex)
        End Try
    End Function
#End Region

#Region "チェック"
    ''' <summary>
    ''' バルブブロックチェック
    ''' </summary>
    ''' <param name="strUseValues"></param>
    ''' <param name="strKataValues"></param>
    ''' <param name="intStart"></param>
    ''' <param name="intEnd"></param>
    ''' <param name="strMsgCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncBlockCheck(strUseValues() As Double, strKataValues() As String, intStart As Integer, _
                                         intEnd As Integer, ByRef strMsgCd As String) As Boolean
        fncBlockCheck = False
        Try
            Dim bolCCheck As Boolean = False
            Dim bolCLCheck As Boolean = False
            Dim bolCDCheck As Boolean = False
            For intRI As Integer = intStart To intEnd
                If strKataValues(intRI).Length > 0 And CInt(strUseValues(intRI)) > 0 Then
                    Dim str() As String = strKataValues(intRI).ToString.Split("-")
                    If str.Length > 0 Then
                        If str(str.Length - 1).StartsWith("CD") Then
                            bolCDCheck = True
                        ElseIf str(str.Length - 1).StartsWith("CL") Then
                            bolCLCheck = True
                        ElseIf str(str.Length - 1).StartsWith("C") And Not str(str.Length - 1).StartsWith("CX") Then
                            bolCCheck = True
                        End If
                    End If
                End If
            Next
            If bolCCheck = True And bolCLCheck = True Then
                strMsgCd = "W1670"
                Exit Function
            End If
            If bolCCheck = True And bolCDCheck = True Then
                strMsgCd = "W8850"
                Exit Function
            End If
            If bolCLCheck = True And bolCDCheck = True Then
                strMsgCd = "W8860"
                Exit Function
            End If
            fncBlockCheck = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' バルブブロックチェック
    ''' </summary>
    ''' <param name="strUseValues"></param>
    ''' <param name="strKataValues"></param>
    ''' <param name="intStart"></param>
    ''' <param name="intEnd"></param>
    ''' <param name="strMsgCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncBlockCheck2(strUseValues() As Double, strKataValues() As String, intStart As Integer, _
                                         intEnd As Integer, ByRef strMsgCd As String) As Boolean
        fncBlockCheck2 = False
        Dim bolCCheck As Boolean = False

        For intRI As Integer = intStart To intEnd
            If strKataValues(intRI).Length > 0 And CInt(strUseValues(intRI)) > 0 Then
                Dim str() As String = strKataValues(intRI).ToString.Split("-")
                If str.Length > 0 Then
                    If str(str.Length - 1).EndsWith("X") Or str(str.Length - 1).EndsWith("X1") _
                        Or str(str.Length - 1).EndsWith("G1") Or str(str.Length - 1).EndsWith("G2") Then
                        bolCCheck = True
                    End If
                End If
            End If
        Next

        If bolCCheck = False Then
            strMsgCd = "W9290"
            Exit Function
        End If

        fncBlockCheck2 = True

    End Function

    ''' <summary>
    ''' バルブブロック使用数チェック
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intStart"></param>
    ''' <param name="intEnd"></param>
    ''' <param name="strMsgCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncMixBlockCheck(objKtbnStrc As KHKtbnStrc, intStart As Integer, _
                                            intEnd As Integer, ByRef strMsgCd As String) As Boolean
        fncMixBlockCheck = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim HTList As New ArrayList
            For intRI As Integer = intStart To intEnd
                '形番要素が選択かつ使用数 > 0の場合
                If strKataValues(intRI).Length > 0 And CInt(strUseValues(intRI)) > 0 Then
                    If strKataValues(intRI).ToString.Contains("-CX") Then
                        If Not HTList.Contains("C4") Then HTList.Add("C4")
                        If Not HTList.Contains("C6") Then HTList.Add("C6")
                    Else
                        Dim str() As String = strKataValues(intRI).ToString.Split("-")
                        If str.Length >= 2 Then
                            If str(str.Length - 2).StartsWith("C") Then
                                If Not HTList.Contains(str(str.Length - 2)) Then
                                    HTList.Add(str(str.Length - 2))
                                End If
                            End If
                        End If
                        If str.Length > 0 Then
                            If str(str.Length - 1).StartsWith("C") Or str(str.Length - 1).StartsWith("M5") Then
                                If Not HTList.Contains(str(str.Length - 1)) Then
                                    HTList.Add(str(str.Length - 1))
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            If HTList.Count <= 1 Then
                strMsgCd = "W1220"
                Exit Function
            End If
            fncMixBlockCheck = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' ミックススイッチチェック
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intStart"></param>
    ''' <param name="intEnd"></param>
    ''' <param name="intHFlag"></param>
    ''' <param name="strMsgCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncMixSwtchCheck(objKtbnStrc As KHKtbnStrc, intStart As Integer, intEnd As Integer, _
                                            intHFlag As Boolean, ByRef strMsgCd As String) As Boolean
        fncMixSwtchCheck = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim HTList As New ArrayList
            For intRI As Integer = intStart To intEnd
                '形番要素が選択かつ使用数 > 0の場合
                If strKataValues(intRI).Length > 0 And CInt(strUseValues(intRI)) > 0 Then
                    If strKataValues(intRI).ToString.Length >= 3 Then
                        Dim strKisyu As String = Left(strKataValues(intRI), 3)
                        If IsNumeric(Left(strKataValues(intRI) & Space(7), 7).Substring(5, 2)) Then
                            Dim strMixBlock As String = Left(strKataValues(intRI) & Space(7), 7).Substring(5, 2)
                            If Not HTList.Contains(strKisyu & "," & strMixBlock) Then
                                HTList.Add(strKisyu & "," & strMixBlock)
                            End If
                        ElseIf strKataValues(intRI).ToString.Contains("-MP") Then
                            If Not HTList.Contains(strKisyu & ",MP") Then
                                HTList.Add(strKisyu & ",MP")
                            End If
                        End If
                    End If
                End If
            Next
            If HTList.Count <= 1 Then
                strMsgCd = "W1190"
                Exit Function
            End If
            '排気誤作動防止弁が"H"の場合
            If intHFlag Then
                Dim myFlag1 As Boolean = False
                Dim myFlag2 As Boolean = False
                'ミックスチェック値(切替位置区分)30,50,MPあれば、他の位置も必要
                For inti As Integer = 0 To HTList.Count - 1
                    If HTList(inti).ToString.Contains(",30") Or _
                        HTList(inti).ToString.Contains(",50") Or _
                        HTList(inti).ToString.Contains(",MP") Then
                        myFlag1 = True
                    Else
                        myFlag2 = True
                    End If
                Next
                If myFlag1 And Not myFlag2 Then
                    strMsgCd = "W1660"
                    Exit Function
                End If
            End If
            fncMixSwtchCheck = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 重複チェック
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intColStart">チェック開始行</param>
    ''' <param name="intColEnd">チェック終了行</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncDblCheck(objKtbnStrc As KHKtbnStrc, ByVal intColStart As Integer, _
                                 ByVal intColEnd As Integer) As Boolean
        fncDblCheck = False
        Try
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            For intCI As Integer = intColStart To intColEnd - 1
                For intCI2 As Integer = intCI + 1 To intColEnd
                    If strKataValues(intCI - 1) <> "" And _
                       strKataValues(intCI2 - 1) <> "" And _
                       strKataValues(intCI - 1) = strKataValues(intCI2 - 1) Then
                        Exit Function
                    End If
                Next
            Next
            fncDblCheck = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 取付レール長さチェック
    ''' </summary>
    ''' <param name="dblRailLen_K"></param>
    ''' <param name="dblRailLen_U"></param>
    ''' <param name="dblStdNum"></param>
    ''' <param name="strMsgCD"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncRailchk(ByVal dblRailLen_K As Double, ByVal dblRailLen_U As Double, ByVal dblStdNum As Double, ByRef strMsgCD As String) As Boolean
        fncRailchk = False
        Try
            If Right(dblRailLen_K, 2) = ".0" Or Right(dblRailLen_K, 2) = ".5" Or InStr(dblRailLen_K, ".") = 0 Then
            Else
                strMsgCD = "W1340"
                Exit Function
            End If
            If dblRailLen_K < (dblStdNum - 25) Or dblRailLen_K > CDbl(32767) Then
                strMsgCD = "W1320"
                Exit Function
            End If

            If Right(dblRailLen_U, 2) = ".0" Or Right(dblRailLen_U, 2) = ".5" Or InStr(dblRailLen_U, ".") = 0 Then
            Else
                strMsgCD = "W1340"
                Exit Function
            End If
            If dblRailLen_U < (dblStdNum - 25) Or dblRailLen_U > CDbl(32767) Then
                strMsgCD = "W1320"
                Exit Function
            End If
            fncRailchk = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' ブランクプラグとサイレンサ検査成績所の使用数チェック
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intStart"></param>
    ''' <param name="intEnd"></param>
    ''' <param name="intRail"></param>
    ''' <param name="strMsgCd"></param>
    ''' <param name="intMax"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncOtherKataCheck(objKtbnStrc As KHKtbnStrc, intStart As Integer, intEnd As Integer, intRail As Integer, _
                                             ByRef strMsgCd As String, Optional intMax As Integer = 1000) As Boolean
        fncOtherKataCheck = False
        For intRI As Integer = intStart - 1 To intEnd - 1
            If intRI <> intRail - 1 Then
                If CInt(objKtbnStrc.strcSelection.intQuantity(intRI)) > 0 And Len(objKtbnStrc.strcSelection.strOptionKataban(intRI)) = 0 Then
                    strMsgCd = "W1310"
                    Exit Function
                End If
                If Int(objKtbnStrc.strcSelection.intQuantity(intRI)) > intMax Then
                    strMsgCd = "W1320"
                    Exit Function
                End If
            End If
        Next
        fncOtherKataCheck = True
    End Function

    ''' <summary>
    ''' 仕様画面でOKボタン押したら、ﾁｪｯｸする（Web系と共通になる）
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="ManifoldMode"></param>
    ''' <param name="dblStdNum"></param>
    ''' <param name="strMsg"></param>
    ''' <param name="strMsgCd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function InputCheck(objKtbnStrc As KHKtbnStrc, ByVal ManifoldMode As Integer, _
                                ByRef dblStdNum As Double, ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        InputCheck = False

        Try
            Dim HT_Option As Hashtable = GetOptionData(objKtbnStrc, ManifoldMode)

            Select Case ManifoldMode
                Case 1
                    If Not ClsInputCheck_01.fncInputChk(objKtbnStrc, HT_Option, dblStdNum, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 2
                    If Not ClsInputCheck_02.fncInputChk(objKtbnStrc, dblStdNum, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 3
                    If Not ClsInputCheck_03.fncInputChk(objKtbnStrc, dblStdNum, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 4
                    If Not ClsInputCheck_04.fncInputChk(objKtbnStrc, HT_Option, dblStdNum, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 5
                    If Not ClsInputCheck_05.fncInpChk(objKtbnStrc, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 6
                    If Not ClsInputCheck_06.fncInpChk(objKtbnStrc, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 7
                    If Not ClsInputCheck_07.fncInputChk(objKtbnStrc, HT_Option, dblStdNum, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 8
                    If Not ClsInputCheck_08.fncInputChk(objKtbnStrc, dblStdNum, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 9
                    If Not ClsInputCheck_09.fncInputChk(objKtbnStrc, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 10
                    If Not ClsInputCheck_10.fncInputChk(objKtbnStrc, dblStdNum, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 11
                    If Not ClsInputCheck_11.fncInputChk(objKtbnStrc, dblStdNum, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 12
                    If Not ClsInputCheck_12.fncInputChk(objKtbnStrc, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 13
                    If Not ClsInputCheck_13.fncInputChk(objKtbnStrc, dblStdNum, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 14
                    If Not ClsInputCheck_14.fncInputChk(objKtbnStrc, dblStdNum, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 15
                    If Not ClsInputCheck_15.fncInputChk(objKtbnStrc, dblStdNum, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 16
                    If Not ClsInputCheck_16.fncInputChk(objKtbnStrc, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 17
                    If Not ClsInputCheck_17.fncInputChk(objKtbnStrc, strMsg, strMsgCd) Then
                        Exit Function
                    End If
                Case 18
                    If Not ClsInputCheck_18.fncInputChk(objKtbnStrc, dblStdNum, strMsg, strMsgCd) Then
                        Exit Function
                    End If
            End Select
            InputCheck = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' データセットから形番と使用数を取得
    ''' </summary>
    ''' <param name="ds"></param>
    ''' <param name="strSelKataban">形番</param>
    ''' <param name="strInpUse">使用数</param>
    ''' <param name="ManifoldMode"></param>
    ''' <remarks></remarks>
    Private Sub subGetInfoFromDS(ByVal ds As DataSet, ByRef strSelKataban As ArrayList, ByRef strInpUse As ArrayList, ByVal ManifoldMode As String)
        '選択した形番と使用数を取得
        Select Case ManifoldMode
            Case 0
                Dim dt As DataTable = ds.Tables("simple")
                For inti As Integer = 0 To dt.Rows.Count - 1
                    If dt.Rows(inti)("ColKata").ToString.Length > 0 Then
                        strSelKataban.Add(dt.Rows(inti)("ColKata").ToString)
                        strSelKataban.Add(dt.Rows(inti)("Col0").ToString)
                    Else
                        strSelKataban.Add(String.Empty)
                    End If
                    If dt.Rows(inti)("Col0").ToString.Length > 0 Then
                        strInpUse.Add(dt.Rows(inti)("Col0").ToString)
                    Else
                        strInpUse.Add(0)
                    End If
                Next
            Case Else
                Dim dt_title As DataTable = ds.Tables("title")
                Dim dt_data As DataTable = ds.Tables("data")

                For inti As Integer = 0 To dt_title.Rows.Count - 1
                    If dt_title.Rows(inti)("ColKata").ToString.Length > 0 Then
                        strSelKataban.Add(dt_title.Rows(inti)("ColKata").ToString)
                    Else
                        strSelKataban.Add(String.Empty)
                    End If
                    If dt_data.Rows(inti)("Col0").ToString.Length > 0 Then
                        strInpUse.Add(dt_data.Rows(inti)("Col0").ToString)
                    Else
                        strInpUse.Add(0)
                    End If
                Next
        End Select
    End Sub
#End Region

    ''' <summary>
    ''' 品名マスタの取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strLanguage"></param>
    ''' <param name="strSpecNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LoadPositionData(ByVal objCon As SqlConnection, strLanguage As String, _
                                      strSpecNo As String) As DataTable
        Dim dtResult As New DataTable
        Dim dalSiyou As New SiyouDAL

        Try
            dtResult = dalSiyou.LoadPositionData(objCon, strLanguage, strSpecNo)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return dtResult
    End Function
    ''' <summary>
    ''' オプションデータの取得
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intMode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function GetOptionData(objKtbnStrc As KHKtbnStrc, intMode As Integer) As Hashtable
        Dim strComma As String = CdCst.Sign.Delimiter.Comma                       'カンマ
        GetOptionData = New Hashtable
        Try
            Select Case intMode
                Case 1
                    Dim strSwitchPos As String = String.Empty
                    Dim strConCaliber As String = String.Empty
                    Dim strMaxSeq As String = String.Empty
                    Dim strOptionT As String = Nothing
                    Dim strOptionD As String = Nothing
                    Dim strOptionP7 As String = String.Empty
                    Dim strOptions() As String = Nothing

                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "MN3Q0", "MT3Q0"
                            strSwitchPos = objKtbnStrc.strcSelection.strOpSymbol(1)
                            strConCaliber = objKtbnStrc.strcSelection.strOpSymbol(3)
                            strMaxSeq = objKtbnStrc.strcSelection.strOpSymbol(7)
                            strOptions = Nothing

                            'オプションＴ保持
                            Dim intFlag As Integer = 0
                            strOptions = objKtbnStrc.strcSelection.strOpSymbol(5).Split(strComma)
                            For intI As Integer = 0 To UBound(strOptions)
                                Select Case strOptions(intI)
                                    Case "T30", "T30R", "T30U", "T30UR", "T51", "T51R", "T51U", "T51UR", _
                                         "T53", "T53R", "T53U", "T53UR", "TX"
                                        If strOptionT Is Nothing Then
                                            strOptionT = strOptions(intI)
                                            intFlag = intFlag + 1
                                        End If
                                End Select
                                If intFlag > 1 Then Exit For
                            Next
                        Case Else
                            If objKtbnStrc.strcSelection.strSeriesKataban = "MN3EX0" Or _
                                objKtbnStrc.strcSelection.strSeriesKataban = "MN4EX0" Then
                                strSwitchPos = ""
                                strConCaliber = objKtbnStrc.strcSelection.strOpSymbol(1)
                                strMaxSeq = objKtbnStrc.strcSelection.strOpSymbol(7)
                            Else
                                strSwitchPos = objKtbnStrc.strcSelection.strOpSymbol(1)
                                strConCaliber = objKtbnStrc.strcSelection.strOpSymbol(3)
                                strMaxSeq = objKtbnStrc.strcSelection.strOpSymbol(9)
                            End If

                            'Ｐ７選択判定
                            If objKtbnStrc.strcSelection.strSeriesKataban = "MN3EX0" Or _
                                objKtbnStrc.strcSelection.strSeriesKataban = "MN4EX0" Then
                                strOptions = objKtbnStrc.strcSelection.strOpSymbol(9).Split(strComma)
                            Else
                                strOptions = objKtbnStrc.strcSelection.strOpSymbol(11).Split(strComma)
                            End If

                            For intI As Integer = 0 To UBound(strOptions)
                                If strOptions(intI).Contains("P7") Then
                                    strOptionP7 = strOptions(intI)
                                    Exit For
                                End If
                            Next
                            strOptions = Nothing

                            'オプションＤ・オプションＴ保持
                            Dim intFlag As Integer = 0
                            If objKtbnStrc.strcSelection.strSeriesKataban = "MN3EX0" Or _
                                objKtbnStrc.strcSelection.strSeriesKataban = "MN4EX0" Then
                                strOptions = objKtbnStrc.strcSelection.strOpSymbol(4).Split(strComma)
                            Else
                                strOptions = objKtbnStrc.strcSelection.strOpSymbol(6).Split(strComma)
                            End If
                            For intI As Integer = 0 To UBound(strOptions)
                                If strOptions(intI).ToString.Length <= 0 Then Continue For
                                Select Case Strings.Left(strOptions(intI), 1)
                                    Case "T"
                                        If strOptionT Is Nothing Then
                                            strOptionT = strOptions(intI)
                                            intFlag = intFlag + 1
                                        End If
                                    Case "D"
                                        If strOptionD Is Nothing Then
                                            strOptionD = strOptions(intI)
                                            intFlag = intFlag + 1
                                        End If
                                End Select
                                If intFlag > 1 Then Exit For
                            Next
                    End Select
                    GetOptionData.Add("strSwitchPos", strSwitchPos)
                    GetOptionData.Add("strConCaliber", strConCaliber)
                    GetOptionData.Add("strMaxSeq", strMaxSeq)
                    GetOptionData.Add("strOptionT", strOptionT)
                    GetOptionData.Add("strOptionD", strOptionD)
                    GetOptionData.Add("strOptionP7", strOptionP7)
                Case 4
                    Dim strOption As String = String.Empty
                    Dim strMountType As String = String.Empty
                    Dim strRensu As String = String.Empty
                    Dim strCleanShiyo As String = String.Empty
                    Dim strSolenoidPos As String                           '切替位置区分
                    Dim strPortSize As String                              '接続口径
                    Dim strElecConType As String                           '電線接続タイプ

                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "R", "U", "S", "V"
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                Case "M3GA1", "M3GB1", "M4GA1", "M4GB1", "M3GA2", "M3GB2", "M4GA2", "M4GB2"
                                    strOption = objKtbnStrc.strcSelection.strOpSymbol(8)           'オプション
                                    strMountType = objKtbnStrc.strcSelection.strOpSymbol(9)        'マウントタイプ
                                    strRensu = objKtbnStrc.strcSelection.strOpSymbol(10)            '連数
                                    strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(12)      'クリーン仕様
                                Case "M3GD1", "M3GE1", "M3GD2", "M3GE2", "M4GD1", "M4GE1", "M4GD2", "M4GE2"
                                    strOption = objKtbnStrc.strcSelection.strOpSymbol(8)           'オプション
                                    strMountType = objKtbnStrc.strcSelection.strOpSymbol(9)        'マウントタイプ
                                    strRensu = objKtbnStrc.strcSelection.strOpSymbol(11)            '連数
                                    strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(13)      'クリーン仕様
                                Case "M3GA3", "M4GA3", "M4GB3"
                                    strOption = objKtbnStrc.strcSelection.strOpSymbol(7)           'オプション
                                    strMountType = objKtbnStrc.strcSelection.strOpSymbol(8)        'マウントタイプ
                                    strRensu = objKtbnStrc.strcSelection.strOpSymbol(9)            '連数
                                    strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(11)      'クリーン仕様
                                Case "M3GD3", "M4GD3", "M4GE3"
                                    strOption = objKtbnStrc.strcSelection.strOpSymbol(7)           'オプション
                                    strMountType = objKtbnStrc.strcSelection.strOpSymbol(8)        'マウントタイプ
                                    strRensu = objKtbnStrc.strcSelection.strOpSymbol(10)            '連数
                                    strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(12)      'クリーン仕様
                            End Select
                            strSolenoidPos = objKtbnStrc.strcSelection.strOpSymbol(1)                           '切替位置区分
                            strPortSize = objKtbnStrc.strcSelection.strOpSymbol(4)                              '接続口径
                            strElecConType = objKtbnStrc.strcSelection.strOpSymbol(5)                           '電線接続タイプ
                        Case Else
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                Case "M3GA1", "M3GB1", "M4GA1", "M4GB1", "M3GA2", "M3GB2", "M4GA2", "M4GB2",
                                     "M3GD1", "M3GE1", "M3GD2", "M3GE2", "M4GD1", "M4GE1", "M4GD2", "M4GE2"
                                    strOption = objKtbnStrc.strcSelection.strOpSymbol(7)           'オプション
                                    strMountType = objKtbnStrc.strcSelection.strOpSymbol(8)        'マウントタイプ
                                    strRensu = objKtbnStrc.strcSelection.strOpSymbol(9)            '連数
                                    strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(11)      'クリーン仕様
                                Case "M3GA3", "M4GA3", "M4GA4", "M4GB3", "M4GB4", "M3GD3", "M4GD3", "M4GE3", "M4GD4", "M4GE4"
                                    strOption = objKtbnStrc.strcSelection.strOpSymbol(6)           'オプション
                                    strMountType = objKtbnStrc.strcSelection.strOpSymbol(7)        'マウントタイプ
                                    strRensu = objKtbnStrc.strcSelection.strOpSymbol(8)            '連数
                                    strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(10)      'クリーン仕様
                            End Select
                            strSolenoidPos = objKtbnStrc.strcSelection.strOpSymbol(1)                           '切替位置区分
                            strPortSize = objKtbnStrc.strcSelection.strOpSymbol(3)                              '接続口径
                            strElecConType = objKtbnStrc.strcSelection.strOpSymbol(4)                           '電線接続タイプ
                    End Select

                    'Select Case objKtbnStrc.strcSelection.strSeriesKataban
                    '    Case "M3GA1", "M3GB1", "M4GA1", "M4GB1", "M3GA2", "M3GB2", "M4GA2", "M4GB2",
                    '         "M3GD1", "M3GE1", "M3GD2", "M3GE2", "M4GD1", "M4GE1", "M4GD2", "M4GE2"
                    '        strOption = objKtbnStrc.strcSelection.strOpSymbol(7)           'オプション
                    '        strMountType = objKtbnStrc.strcSelection.strOpSymbol(8)        'マウントタイプ
                    '        strRensu = objKtbnStrc.strcSelection.strOpSymbol(9)            '連数
                    '        strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(11)      'クリーン仕様
                    '    Case "M3GA3", "M4GA3", "M4GA4", "M4GB3", "M4GB4", "M3GD3", "M4GD3", "M4GE3", "M4GD4", "M4GE4"
                    '        strOption = objKtbnStrc.strcSelection.strOpSymbol(6)           'オプション
                    '        strMountType = objKtbnStrc.strcSelection.strOpSymbol(7)        'マウントタイプ
                    '        strRensu = objKtbnStrc.strcSelection.strOpSymbol(8)            '連数
                    '        strCleanShiyo = objKtbnStrc.strcSelection.strOpSymbol(10)      'クリーン仕様
                    'End Select
                    'Dim strSolenoidPos As String = objKtbnStrc.strcSelection.strOpSymbol(1)    '切替位置区分
                    'Dim strPortSize As String = objKtbnStrc.strcSelection.strOpSymbol(3)       '接続口径
                    'Dim strElecConType As String = objKtbnStrc.strcSelection.strOpSymbol(4)    '電線接続タイプ
                    Dim strMaxSeq As String = strRensu                                         '連数
                    Dim strSolenoidType As String = String.Empty
                    Dim strOptionK As String = String.Empty
                    Dim strStdMFType As String = String.Empty
                    Dim strOptionH As String = String.Empty
                    Dim strOptionZ1 As String = String.Empty
                    Dim strOptionZ2 As String = String.Empty
                    Dim strOptionZ3 As String = String.Empty
                    Dim strOptionD As String = String.Empty
                    Dim strOptionP70 As String = String.Empty
                    Dim strOptionX As String = String.Empty
                    Dim strOptionG As String = String.Empty

                    '操作区分
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2)
                        Case "0"
                            strSolenoidType = "9"  '操作区分（ﾊﾟﾌﾞﾘｯｸ変数）
                            'マスターバルブ
                        Case "1"
                            strSolenoidType = "8"  '操作区分（ﾊﾟﾌﾞﾘｯｸ変数）
                    End Select

                    'オプション選択判定
                    Dim strOptions() As String = strOption.Split(strComma)
                    For intI As Integer = 0 To UBound(strOptions)
                        If strOptions(intI).Contains("K") Then
                            strOptionK = strOptions(intI)
                            strStdMFType = "P"
                        ElseIf strOptions(intI).Contains("H") Then
                            strOptionH = strOptions(intI)
                        ElseIf strOptions(intI).Contains("Z1") Then
                            strOptionZ1 = strOptions(intI)
                        ElseIf strOptions(intI).Contains("Z2") Then
                            strOptionZ2 = strOptions(intI)
                        ElseIf strOptions(intI).Contains("Z3") Then
                            strOptionZ3 = strOptions(intI)
                        ElseIf strOptions(intI).Contains("X") Or strOptions(intI).Contains("X1") Then
                            strOptionX = strOptions(intI)
                        ElseIf strOptions(intI).Contains("G1") Or strOptions(intI).Contains("G2") Then
                            strOptionG = strOptions(intI)
                        End If
                    Next
                    If strMountType.Trim = "D" Then strOptionD = "D"
                    If strCleanShiyo.Trim = "P70" Then strOptionP70 = "P70"

                    GetOptionData.Add("strSolenoidPos", strSolenoidPos)
                    GetOptionData.Add("strPortSize", strPortSize)
                    GetOptionData.Add("strElecConType", strElecConType)
                    GetOptionData.Add("strMaxSeq", strMaxSeq)
                    GetOptionData.Add("strSolenoidType", strSolenoidType)
                    GetOptionData.Add("strOptionK", strOptionK)
                    GetOptionData.Add("strStdMFType", strStdMFType)
                    GetOptionData.Add("strOptionH", strOptionH)
                    GetOptionData.Add("strOptionZ1", strOptionZ1)
                    GetOptionData.Add("strOptionZ2", strOptionZ2)
                    GetOptionData.Add("strOptionZ3", strOptionZ3)
                    GetOptionData.Add("strOptionD", strOptionD)
                    GetOptionData.Add("strOptionP70", strOptionP70)
                    GetOptionData.Add("strOptionX", strOptionX)
                    GetOptionData.Add("strOptionG", strOptionG)

                Case 7
                    Dim strOption As String = String.Empty
                    Dim strOptionX As String = String.Empty
                    Dim strOptionG As String = String.Empty

                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "MN4GB1", "MN4GB2", "MN4GBX12"
                            strOption = objKtbnStrc.strcSelection.strOpSymbol(8)           'オプション
                    End Select

                    'オプション選択判定
                    Dim strOptions() As String = strOption.Split(strComma)
                    For intI As Integer = 0 To UBound(strOptions)
                        If strOptions(intI).Contains("X") Or strOptions(intI).Contains("X1") Then
                            strOptionX = strOptions(intI)
                        ElseIf strOptions(intI).Contains("G1") Or strOptions(intI).Contains("G2") Then
                            strOptionG = strOptions(intI)
                        End If
                    Next

                    GetOptionData.Add("strOptionX", strOptionX)
                    GetOptionData.Add("strOptionG", strOptionG)

            End Select
        Catch ex As Exception
            Call WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' CX選択肢の取得
    ''' </summary>
    ''' <param name="strMode"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strKata"></param>
    ''' <param name="intRow"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCXList(strMode As String, objKtbnStrc As KHKtbnStrc, strKata As String, intRow As Integer) As ArrayList
        GetCXList = New ArrayList
        Try
            Dim strCX As String = String.Empty
            Dim strKey As String = String.Empty
            Select Case strMode
                Case "3"
                    If intRow >= CdCst.Siyou_03.Elect1 - 1 And intRow <= CdCst.Siyou_03.Elect14 - 1 Then
                        strKey = objKtbnStrc.strcSelection.strOpSymbol(6).ToString
                        If strKata.EndsWith("-CX") Then
                            If strKey = "CXF" Then
                                strCX = "C4F,C6F"
                            ElseIf strKey = "CX" Then
                                strCX = "C4,C6"
                            End If
                            If objKtbnStrc.strcSelection.strKeyKataban = "3" Or _
                                objKtbnStrc.strcSelection.strKeyKataban = "4" Then
                                strCX &= "X"
                            End If
                        End If
                    End If
                Case "4"
                    If intRow >= CdCst.Siyou_04.Valve1 - 1 And intRow <= CdCst.Siyou_04.Spacer4 - 1 Then
                        If strKata.Contains("-CX") Then

                            'CX継手区分の取得
                            Dim intCXKbn As Integer = fncGetCXKbn(objKtbnStrc.strcSelection.strSeriesKataban, objKtbnStrc.strcSelection.strKeyKataban)

                            strKey = objKtbnStrc.strcSelection.strOpSymbol(2).ToString '操作区分

                            Select Case Strings.Left(strKata, 4)
                                Case "3GB1", "3GE1"

                                    Select Case intCXKbn
                                        Case 1
                                            strCX = "C4,C6,X"
                                        Case 2
                                            strCX = "C6,C8,X"
                                        Case 3
                                            strCX = "C8,C10,X"
                                        Case Else
                                            If strKey = "0" Then   '操作区分
                                                strCX = "C4,C6,C18,CD4,CD6,X"
                                            Else
                                                strCX = "C4,C6,X"
                                            End If
                                    End Select

                                Case "3GB2"

                                    Select Case intCXKbn
                                        Case 1
                                            strCX = "C4,C6,X"
                                        Case 2
                                            strCX = "C6,C8,X"
                                        Case 3
                                            strCX = "C8,C10,X"
                                        Case Else
                                            If strKey = "0" Then   '操作区分
                                                strCX = "C4,C6,C8,CD6,CD8,X"
                                            Else
                                                strCX = "C4,C6,C8,X"
                                            End If
                                    End Select

                                Case "4GB1", "4GE1"
                                    If strKata.StartsWith("4GB11") Or strKata.StartsWith("4GE11") Then

                                        Select Case intCXKbn
                                            Case 1
                                                strCX = "C4,C6,X"
                                            Case 2
                                                strCX = "C6,C8,X"
                                            Case 3
                                                strCX = "C8,C10,X"
                                            Case Else
                                                If strKey = "0" Then   '操作区分
                                                    strCX = "C4,C6,C18,CD4,CD6,CF,CL4,CL6,X"
                                                Else
                                                    strCX = "C4,C6,CL4,CL6,X"
                                                End If
                                        End Select

                                    Else
                                        Select Case intCXKbn
                                            Case 1
                                                strCX = "C4,C6,X"
                                            Case 2
                                                strCX = "C6,C8,X"
                                            Case 3
                                                strCX = "C8,C10,X"
                                            Case Else
                                                If strKey = "0" Then   '操作区分
                                                    strCX = "C4,C6,C18,CD4,CD6,CF,X"
                                                Else
                                                    strCX = "C4,C6,X"
                                                End If
                                        End Select
                                    End If

                                    '調査必要（要るかどうか）
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                                        Case "C"
                                            strCX = "C4,C6,X"
                                    End Select

                                Case "4GB2"

                                    Select Case intCXKbn
                                        Case 1
                                            strCX = "C4,C6,X"
                                        Case 2
                                            strCX = "C6,C8,X"
                                        Case 3
                                            strCX = "C8,C10,X"
                                        Case Else
                                            If strKata.StartsWith("4GB21") Then
                                                If strKey = "0" Then   '操作区分
                                                    strCX = "C4,C6,C8,CD6,CD8,CL6,CL8,X"
                                                Else
                                                    strCX = "C4,C6,C8,CL6,CL8,X"
                                                End If
                                            Else
                                                If strKey = "0" Then   '操作区分
                                                    strCX = "C4,C6,C8,CD6,CD8,X"
                                                Else
                                                    strCX = "C4,C6,C8,X"
                                                End If
                                            End If
                                    End Select

                                Case "4GB3"
                                    Select Case intCXKbn
                                        Case 1
                                            strCX = "C4,C6,X"
                                        Case 2
                                            strCX = "C6,C8,X"
                                        Case 3
                                            'C6追加により、ここにC6を追加する  2017/04/11 追加
                                            strCX = "C6,C8,C10,X"
                                            'strCX = "C8,C10,X"
                                        Case Else
                                            If strKata.StartsWith("4GB31") Then
                                                If strKey = "0" Then   '操作区分
                                                    strCX = "C6,C8,C10,CD8,CD10,CL8,CL10,X"
                                                Else
                                                    strCX = "C6,C8,C10,CL8,CL10,X"
                                                End If
                                            Else
                                                If strKey = "0" Then   '操作区分
                                                    strCX = "C6,C8,C10,CD8,CD10,X"
                                                Else
                                                    strCX = "C6,C8,C10,X"
                                                End If
                                            End If
                                    End Select
                                Case "4GB4", "4GE4"
                                    Select Case intCXKbn
                                        Case 1
                                            strCX = "C4,C6,X"
                                        Case 2
                                            strCX = "C6,C8,X"
                                        Case 3
                                            strCX = "C8,C10,X"
                                        Case Else
                                            strCX = "C8,C10,C12"
                                    End Select
                                Case "3GE2"
                                    Select Case intCXKbn
                                        Case 1
                                            strCX = "C4,C6,X"
                                        Case 2
                                            strCX = "C6,C8,X"
                                        Case 3
                                            strCX = "C8,C10,X"
                                        Case Else
                                            strCX = "C4,C6,C8,X"
                                    End Select
                                Case "4GE2"
                                    Select Case intCXKbn
                                        Case 1
                                            strCX = "C4,C6,X"
                                        Case 2
                                            strCX = "C6,C8,X"
                                        Case 3
                                            strCX = "C8,C10,X"
                                        Case Else
                                            If strKata.StartsWith("4GE21") Then
                                                strCX = "C4,C6,C8,CL6,CL8,X"
                                            Else
                                                strCX = "C4,C6,C8,X"
                                            End If
                                    End Select
                                Case "4GE3"
                                    Select Case intCXKbn
                                        Case 1
                                            strCX = "C4,C6,X"
                                        Case 2
                                            strCX = "C6,C8,X"
                                        Case 3
                                            strCX = "C8,C10,X"
                                        Case Else
                                            If strKata.StartsWith("4GE31") Then
                                                strCX = "C6,C8,C10,CL8,CL10,X"
                                            Else
                                                strCX = "C6,C8,C10,X"
                                            End If
                                    End Select
                            End Select
                            If strKata.StartsWith("4G1-MP-") Or strKata.StartsWith("4G1R-MP-") Then
                                Select Case intCXKbn
                                    Case 1
                                        strCX = "C4,C6,X"
                                    Case 2
                                        strCX = "C6,C8,X"
                                    Case 3
                                        'C6追加により、ここにC6を追加する  2017/04/11 追加
                                        strCX = "C6,C8,C10,X"
                                        'strCX = "C8,C10,X"
                                    Case Else
                                        If strKey = "0" Then   '操作区分
                                            strCX = "C4,C6,C18,CD4,CD6,CF,CL4,CL6,X"
                                        Else
                                            strCX = "C4,C6,CL4,CL6,X"
                                        End If
                                End Select
                            ElseIf strKata.StartsWith("4G2-MP-") Or strKata.StartsWith("4G2R-MP-") Then
                                Select Case intCXKbn
                                    Case 1
                                        strCX = "C4,C6,X"
                                    Case 2
                                        strCX = "C6,C8,X"
                                    Case 3
                                        'C6追加により、ここにC6を追加する  2017/04/11 追加
                                        strCX = "C6,C8,C10,X"
                                        'strCX = "C8,C10,X"
                                    Case Else
                                        If strKey = "0" Then   '操作区分
                                            strCX = "C4,C6,C8,CD6,CD8,CL6,CL8,X"
                                        Else
                                            strCX = "C4,C6,C8,CL6,CL8,X"
                                        End If
                                End Select
                            ElseIf strKata.StartsWith("4G3-MP-") Or strKata.StartsWith("4G3R-MP-") Then
                                Select Case intCXKbn
                                    Case 1
                                        strCX = "C4,C6,X"
                                    Case 2
                                        strCX = "C6,C8,X"
                                    Case 3
                                        'C6追加により、ここにC6を追加する  2017/04/11 追加
                                        strCX = "C6,C8,C10,X"
                                        'strCX = "C8,C10,X"
                                    Case Else
                                        If strKey = "0" Then   '操作区分
                                            strCX = "C6,C8,C10,CD8,CD10,CL8,CL10,X"
                                        Else
                                            strCX = "C6,C8,C10,CL8,CL10,X"
                                        End If
                                End Select
                            ElseIf strKata.StartsWith("4G1-MPS-") Or strKata.StartsWith("4G1R-MPS-") Then
                                Select Case intCXKbn
                                    Case 1
                                        strCX = "C4,C6,X"
                                    Case 2
                                        strCX = "C6,C8,X"
                                    Case 3
                                        strCX = "C8,C10,X"
                                    Case Else
                                        If strKey = "0" Then   '操作区分
                                            strCX = "C4,C6,C18,CD4,CD6,CF,CL4,CL6,X"
                                        Else
                                            strCX = "C4,C6,CL4,CL6,X"
                                        End If
                                End Select
                            ElseIf strKata.StartsWith("4G1-MPD-") Or strKata.StartsWith("4G1R-MPD-") Then
                                Select Case intCXKbn
                                    Case 1
                                        strCX = "C4,C6,X"
                                    Case 2
                                        strCX = "C6,C8,X"
                                    Case 3
                                        strCX = "C8,C10,X"
                                    Case Else
                                        If strKey = "0" Then   '操作区分
                                            strCX = "C4,C6,C18,CD4,CD6,CF,X"
                                        Else
                                            strCX = "C4,C6,X"
                                        End If
                                End Select
                            ElseIf strKata.StartsWith("4G2-MPS-") Or strKata.StartsWith("4G2R-MPS-") Then
                                Select Case intCXKbn
                                    Case 1
                                        strCX = "C4,C6,X"
                                    Case 2
                                        strCX = "C6,C8,X"
                                    Case 3
                                        strCX = "C8,C10,X"
                                    Case Else
                                        If strKey = "0" Then   '操作区分
                                            strCX = "C4,C6,C8,CD6,CD8,CL6,CL8,X"
                                        Else
                                            strCX = "C4,C6,C8,CL6,CL8,X"
                                        End If
                                End Select
                            ElseIf strKata.StartsWith("4G2-MPD-") Or strKata.StartsWith("4G2R-MPD-") Then
                                Select Case intCXKbn
                                    Case 1
                                        strCX = "C4,C6,X"
                                    Case 2
                                        strCX = "C6,C8,X"
                                    Case 3
                                        strCX = "C8,C10,X"
                                    Case Else
                                        If strKey = "0" Then   '操作区分
                                            strCX = "C4,C6,C8,CD6,CD8,X"
                                        Else
                                            strCX = "C4,C6,C8,X"
                                        End If
                                End Select
                            ElseIf strKata.StartsWith("4G3-MPS-") Or strKata.StartsWith("4G3R-MPS-") Then
                                Select Case intCXKbn
                                    Case 1
                                        strCX = "C4,C6,X"
                                    Case 2
                                        strCX = "C6,C8,X"
                                    Case 3
                                        'C6追加により、ここにC6を追加する  2017/04/11 追加
                                        strCX = "C6,C8,C10,X"
                                        'strCX = "C8,C10,X"
                                    Case Else
                                        If strKey = "0" Then   '操作区分
                                            strCX = "C6,C8,C10,CD8,CD10,CL8,CL10,X"
                                        Else
                                            strCX = "C6,C8,C10,CL8,CL10,X"
                                        End If
                                End Select
                            ElseIf strKata.StartsWith("4G3-MPD-") Or strKata.StartsWith("4G3R-MPD-") Then
                                Select Case intCXKbn
                                    Case 1
                                        strCX = "C4,C6,X"
                                    Case 2
                                        strCX = "C6,C8,X"
                                    Case 3
                                        'C6追加により、ここにC6を追加する  2017/04/11 追加
                                        strCX = "C6,C8,C10,X"
                                        'strCX = "C8,C10,X"
                                    Case Else
                                        If strKey = "0" Then   '操作区分
                                            strCX = "C6,C8,C10,CD8,CD10,X"
                                        Else
                                            strCX = "C6,C8,C10,X"
                                        End If
                                End Select
                            End If
                        End If
                    End If
            End Select
            If strCX.Length > 0 Then strCX = "," & strCX

            Dim strCXKey() As String = strCX.Split(",")
            For inti As Integer = 0 To strCXKey.Length - 1
                GetCXList.Add(strCXKey(inti).ToString)
            Next
        Catch ex As Exception
            Call WriteErrorLog("E001", ex)
        End Try
    End Function

    ''' <summary>
    ''' 中間行の取得
    ''' </summary>
    ''' <param name="strSpecNo"></param>
    ''' <param name="intRowIdx"></param>
    ''' <param name="bolMain"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetMidRow(strSpecNo As String, intRowIdx As Integer, ByRef bolMain As Boolean) As Boolean
        GetMidRow = False
        bolMain = False
        Select Case strSpecNo
            Case "05"
                Select Case intRowIdx
                    Case CdCst.Siyou_05.ExpCovRep - 1, CdCst.Siyou_05.ExpCovExh - 1
                        If intRowIdx = CdCst.Siyou_05.ExpCovRep - 1 Then bolMain = True
                        Return True
                End Select
            Case "06"
                Select Case intRowIdx
                    Case CdCst.Siyou_06.ExpCovRep - 1, CdCst.Siyou_06.ExpCovExh - 1
                        If intRowIdx = CdCst.Siyou_06.ExpCovRep - 1 Then bolMain = True
                        Return True
                End Select
            Case "09"
                Select Case intRowIdx
                    Case CdCst.Siyou_09.PartitionS - 1, CdCst.Siyou_09.PartitionE - 1
                        If intRowIdx = CdCst.Siyou_09.PartitionS - 1 Then bolMain = True
                        Return True
                End Select
            Case "13"
                Select Case intRowIdx
                    Case CdCst.Siyou_13.Partition1 - 1, CdCst.Siyou_13.Partition2 - 1
                        If intRowIdx = CdCst.Siyou_13.Partition1 - 1 Then bolMain = True
                        Return True
                End Select
            Case "16"
                Select Case intRowIdx
                    Case CdCst.Siyou_16.Partition1 - 1, CdCst.Siyou_16.Partition2 - 1
                        If intRowIdx = CdCst.Siyou_16.Partition1 - 1 Then bolMain = True
                        Return True
                End Select
        End Select
    End Function

    ''' <summary>
    ''' 仕様書情報編集(簡易マニホールドの形番変換)
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="strUpdKigou"></param>
    ''' <remarks></remarks>
    Public Shared Sub subEditSpecInfo_00(ByRef objKtbnStrc As KHKtbnStrc, strUpdKigou() As String)

        Dim strKataban As String = String.Empty
        'Dim strMP As String = "ﾏｽｷﾝｸﾞﾌﾟﾚｰﾄ"

        '属性リストを取得する
        Dim arr_zokusei As New ArrayList
        Dim dt_zokusei As New DS_100Test.kh_item_mstDataTable
        Using da_zokusei As New DS_100TestTableAdapters.kh_item_mstTableAdapter
            da_zokusei.FillbySpec(dt_zokusei, objKtbnStrc.strcSelection.strSpecNo.ToString.Trim)
        End Using
        For inti As Integer = 0 To dt_zokusei.Rows.Count - 1
            If Not dt_zokusei.Rows(inti) Is Nothing AndAlso dt_zokusei.Rows(inti)("item_num").ToString.Length > 0 Then
                For intj As Integer = 0 To CLng(dt_zokusei.Rows(inti)("item_num").ToString) - 1
                    arr_zokusei.Add(dt_zokusei.Rows(inti)("zokusei_cd").ToString)
                Next
            End If
        Next

        Try
            Select Case objKtbnStrc.strcSelection.strSpecNo.ToString.Trim
                Case "51"
                    For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Count - 1
                        If objKtbnStrc.strcSelection.strOptionKataban(idx) = String.Empty And idx = 0 Then Continue For
                        If arr_zokusei(idx) = "D9" Then
                            objKtbnStrc.strcSelection.strOptionKataban(idx) = _
                                Strings.Left(objKtbnStrc.strcSelection.strSeriesKataban, 1) & _
                                objKtbnStrc.strcSelection.strOpSymbol(2).ToString & "-MP"
                        Else
                            objKtbnStrc.strcSelection.strOptionKataban(idx) = _
                                Strings.Left(objKtbnStrc.strcSelection.strOptionKataban(idx).ToString.PadRight(4), 4).Trim & _
                                objKtbnStrc.strcSelection.strOpSymbol(4).ToString
                        End If
                    Next
                Case "52", "60", "61", "62", "63", "65", "67", "69", "71"
                    For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Count - 1
                        If Strings.Left(objKtbnStrc.strcSelection.strOptionKataban(idx), 1) = "A" Then
                        Else
                            If arr_zokusei(idx) = "D9" Then
                                objKtbnStrc.strcSelection.strOptionKataban(idx) = _
                                    Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 3) & "MP"
                            Else
                                objKtbnStrc.strcSelection.strOptionKataban(idx) = _
                                    Strings.Left(objKtbnStrc.strcSelection.strOptionKataban(idx), 4) & "0"
                            End If
                        End If
                    Next
                Case "64", "66", "68", "70", "72", "A4", "A5", "A6", "A7", "A8"
                    For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Count - 1
                        If arr_zokusei(idx) = "D9" Then
                            objKtbnStrc.strcSelection.strOptionKataban(idx) = _
                                Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 3) & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "MP"
                        End If
                    Next
                Case "89", "90"
                    For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Count - 1
                        If arr_zokusei(idx) = "D9" Then
                            objKtbnStrc.strcSelection.strOptionKataban(idx) = _
                                Strings.Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & "-MP"
                        Else
                            objKtbnStrc.strcSelection.strOptionKataban(idx) = _
                                Strings.Left(objKtbnStrc.strcSelection.strOptionKataban(idx), 4) & "9"
                        End If
                    Next
                Case "98"
                    For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Count - 1
                        If arr_zokusei(idx) = "D9" Then
                            objKtbnStrc.strcSelection.strOptionKataban(idx) = _
                                Strings.Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & "-MP"
                        Else
                            objKtbnStrc.strcSelection.strOptionKataban(idx) = _
                                Strings.Left(objKtbnStrc.strcSelection.strOptionKataban(idx), 5) & "9"
                        End If
                    Next
                Case "54", "55", "56", "57", "58", "59", "91", "92"
                    For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Count - 1
                        If arr_zokusei(idx) = "D9" Then
                            objKtbnStrc.strcSelection.strOptionKataban(idx) = _
                                Strings.Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & "-MP"
                        Else
                            objKtbnStrc.strcSelection.strOptionKataban(idx) = _
                                Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 4) & _
                                Mid(objKtbnStrc.strcSelection.strOptionKataban(idx), 5, 1) & "9"
                            If objKtbnStrc.strcSelection.strSeriesKataban = "M3MA0" Then
                                objKtbnStrc.strcSelection.strOptionKataban(idx) = _
                                    objKtbnStrc.strcSelection.strOptionKataban(idx) & "-" & _
                                    objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            End If
                        End If
                    Next
                Case "A1", "A2", "B2", "B3", "B4"
                    'そのまま
                Case "53", "93", "73", "74", "75", "76", "77", "78", "79", _
                    "80", "81", "82", "83", "84", "85", "86", "87", "88"
                    objKtbnStrc.strcSelection.strOptionKataban = strUpdKigou
                Case "S", "T", "U"      'RM1805001_4Rシリーズ追加
                    If objKtbnStrc.strcSelection.strSeriesKataban = "M4HA1" Then
                        For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Count - 1
                            If dt_zokusei.Rows(idx).Item("label_content") = "MP" Then
                                objKtbnStrc.strcSelection.strOptionKataban(idx) = objKtbnStrc.strcSelection.strSeriesKataban & "-MP"
                            End If
                        Next
                    Else
                        For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Count - 1
                            If dt_zokusei.Rows(idx).Item("label_content") = "MP" Then
                                objKtbnStrc.strcSelection.strOptionKataban(idx) = objKtbnStrc.strcSelection.strSeriesKataban & "-MP"
                            End If
                        Next
                    End If
            End Select
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 引当仕様書情報編集(CMF,GMFシリーズ)
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intMode"></param>
    ''' <remarks></remarks>
    Public Shared Sub subEditSpecInfoGMF(ByRef objKtbnStrc As KHKtbnStrc, intMode As Integer)
        Dim strCalibMix As String = ""
        Dim strKataban As String = ""
        Dim strPosInfo As String = ""
        Dim intType1 As Integer
        Dim intType2 As Integer
        Dim intCount As Integer
        Dim sbCoordinates As New StringBuilder
        Dim strKeyKata As String = objKtbnStrc.strcSelection.strKeyKataban
        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban

        Try
            Dim strUpdUse(UBound(objKtbnStrc.strcSelection.intQuantity) + 2) As Double
            Dim strUpdKataban(UBound(objKtbnStrc.strcSelection.strOptionKataban) + 2) As String
            'CHANGED BY YGY 20141119    CMFZ2-HX3-HY3BDU  仕様書出力
            'Dim strUpdPosition(UBound(objKtbnStrc.strcSelection.strPositionInfo) + 1) As String
            Dim strUpdPosition(objKtbnStrc.strcSelection.strPositionInfo.Count + 1) As String

            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            'ABポートブラグ位置設定
            For intI As Integer = 0 To 9
                For intJ As Integer = Siyou_05.ElType1 To Siyou_05.ElType6
                    If arySelectInf(intJ - 1)(intI) = "1" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case ""
                                arySelectInf(Siyou_05.ABPlugL - 1) = StrNewString(arySelectInf(Siyou_05.ABPlugL - 1), intI, "1")
                                strUseValues(Siyou_05.ABPlugL - 1) = CInt(strUseValues(Siyou_05.ABPlugL - 1)) + 1
                            Case "L"
                            Case "H"
                                arySelectInf(Siyou_05.ABPlugR - 1) = StrNewString(arySelectInf(Siyou_05.ABPlugR - 1), intI, "1")
                                strUseValues(Siyou_05.ABPlugR - 1) = CInt(strUseValues(Siyou_05.ABPlugR - 1)) + 1
                            Case "Z"
                                arySelectInf(Siyou_05.ABPlugL - 1) = StrNewString(arySelectInf(Siyou_05.ABPlugL - 1), intI, "1")
                                arySelectInf(Siyou_05.ABPlugR - 1) = StrNewString(arySelectInf(Siyou_05.ABPlugR - 1), intI, "1")
                                strUseValues(Siyou_05.ABPlugL - 1) = CInt(strUseValues(Siyou_05.ABPlugL - 1)) + 1
                                strUseValues(Siyou_05.ABPlugR - 1) = CInt(strUseValues(Siyou_05.ABPlugR - 1)) + 1
                            Case "T"
                        End Select
                    End If
                Next
            Next

            'ABポート接続口径設定
            For intI As Integer = 0 To 9
                For intJ As Integer = Siyou_05.ElType1 To Siyou_05.ElType6
                    If arySelectInf(intJ - 1)(intI) = "1" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Case "02"
                                arySelectInf(Siyou_05.ABCon02 - 1) = StrNewString(arySelectInf(Siyou_05.ABCon02 - 1), intI, "1")
                                strUseValues(Siyou_05.ABCon02 - 1) = CInt(strUseValues(Siyou_05.ABCon02 - 1)) + 1
                            Case "03"
                                arySelectInf(Siyou_05.ABCon03 - 1) = StrNewString(arySelectInf(Siyou_05.ABCon03 - 1), intI, "1")
                                strUseValues(Siyou_05.ABCon03 - 1) = CInt(strUseValues(Siyou_05.ABCon03 - 1)) + 1
                            Case "04"
                                arySelectInf(Siyou_05.ABCon04 - 1) = StrNewString(arySelectInf(Siyou_05.ABCon04 - 1), intI, "1")
                                strUseValues(Siyou_05.ABCon04 - 1) = CInt(strUseValues(Siyou_05.ABCon04 - 1)) + 1
                            Case "HX1"
                            Case "HX2"
                            Case "HX3"
                                Select Case True
                                    Case Left(strKataValues(intJ - 1), 5) = "PV5-6" Or _
                                         Left(strKataValues(intJ - 1), 6) = "PV5G-6" Or _
                                         Left(strKataValues(intJ - 1), 3) = "CM1"
                                        arySelectInf(Siyou_05.ABCon02 - 1) = StrNewString(arySelectInf(Siyou_05.ABCon02 - 1), intI, "1")
                                        strUseValues(Siyou_05.ABCon02 - 1) = CInt(strUseValues(Siyou_05.ABCon02 - 1)) + 1
                                    Case Else
                                        arySelectInf(Siyou_05.ABCon03 - 1) = StrNewString(arySelectInf(Siyou_05.ABCon03 - 1), intI, "1")
                                        strUseValues(Siyou_05.ABCon03 - 1) = CInt(strUseValues(Siyou_05.ABCon03 - 1)) + 1
                                End Select
                            Case "HX4"
                                Select Case True
                                    Case Left(strKataValues(intJ - 1), 5) = "PV5-6" Or _
                                         Left(strKataValues(intJ - 1), 6) = "PV5G-6" Or _
                                         Left(strKataValues(intJ - 1), 3) = "CM1"
                                        arySelectInf(Siyou_05.ABCon02 - 1) = StrNewString(arySelectInf(Siyou_05.ABCon02 - 1), intI, "1")
                                        strUseValues(Siyou_05.ABCon02 - 1) = CInt(strUseValues(Siyou_05.ABCon02 - 1)) + 1
                                    Case Else
                                        arySelectInf(Siyou_05.ABCon04 - 1) = StrNewString(arySelectInf(Siyou_05.ABCon04 - 1), intI, "1")
                                        strUseValues(Siyou_05.ABCon04 - 1) = CInt(strUseValues(Siyou_05.ABCon04 - 1)) + 1
                                End Select
                            Case "HX5"
                                arySelectInf(Siyou_05.ABCon03 - 1) = StrNewString(arySelectInf(Siyou_05.ABCon03 - 1), intI, "1")
                                strUseValues(Siyou_05.ABCon03 - 1) = CInt(strUseValues(Siyou_05.ABCon03 - 1)) + 1
                            Case "HX6"
                                Select Case True
                                    Case Left(strKataValues(intJ - 1), 5) = "PV5-6" Or _
                                         Left(strKataValues(intJ - 1), 6) = "PV5G-6" Or _
                                         Left(strKataValues(intJ - 1), 3) = "CM1"
                                        arySelectInf(Siyou_05.ABCon03 - 1) = StrNewString(arySelectInf(Siyou_05.ABCon03 - 1), intI, "1")
                                        strUseValues(Siyou_05.ABCon03 - 1) = CInt(strUseValues(Siyou_05.ABCon03 - 1)) + 1
                                    Case Else
                                        arySelectInf(Siyou_05.ABCon04 - 1) = StrNewString(arySelectInf(Siyou_05.ABCon04 - 1), intI, "1")
                                        strUseValues(Siyou_05.ABCon04 - 1) = CInt(strUseValues(Siyou_05.ABCon04 - 1)) + 1
                                End Select
                        End Select
                    End If
                Next
            Next

            If Left(objKtbnStrc.strcSelection.strOpSymbol(3), 2) = "HX" Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(3)
                    Case "HX1"
                        intType1 = Siyou_05.ABCon02
                        intType2 = Siyou_05.ABCon03
                    Case "HX2"
                        intType1 = Siyou_05.ABCon03
                        intType2 = Siyou_05.ABCon04
                    Case "HX3"
                        intType1 = Siyou_05.ABCon02
                        intType2 = Siyou_05.ABCon03
                    Case "HX4"
                        intType1 = Siyou_05.ABCon02
                        intType2 = Siyou_05.ABCon04
                    Case "HX5"
                        intType1 = Siyou_05.ElType1
                        intType2 = Siyou_05.ElType2
                    Case "HX6"
                        intType1 = Siyou_05.ABCon03
                        intType2 = Siyou_05.ABCon04
                End Select

                strCalibMix = "-" & CStr(strUseValues(intType1 - 1)) & CStr(strUseValues(intType2 - 1))
            Else
                strCalibMix = ""
            End If

            '↓RM1303003 2013/03/15
            If strSeriesKata.StartsWith("GMF") Then
                Select Case strKeyKata
                    Case "1", "2", "3"
                        strKataban = strSeriesKata & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                            objKtbnStrc.strcSelection.strOpSymbol(2) & "-" & _
                            objKtbnStrc.strcSelection.strOpSymbol(3) & _
                            objKtbnStrc.strcSelection.strOpSymbol(4) & "-" & _
                            objKtbnStrc.strcSelection.strOpSymbol(5) & _
                            objKtbnStrc.strcSelection.strOpSymbol(6) & _
                            objKtbnStrc.strcSelection.strOpSymbol(7) & _
                            objKtbnStrc.strcSelection.strOpSymbol(8)
                        If objKtbnStrc.strcSelection.strOpSymbol(9) = "" And objKtbnStrc.strcSelection.strOpSymbol(10) = "" Then
                            strKataban &= strCalibMix
                        Else
                            strKataban &= "-" & objKtbnStrc.strcSelection.strOpSymbol(9) & _
                                objKtbnStrc.strcSelection.strOpSymbol(10) & strCalibMix
                        End If
                End Select
            Else
                Select Case strKeyKata
                    Case "1", "2", "3", "8"
                        strKataban = strSeriesKata & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                            objKtbnStrc.strcSelection.strOpSymbol(2) & "-" & _
                            objKtbnStrc.strcSelection.strOpSymbol(3) & _
                            objKtbnStrc.strcSelection.strOpSymbol(4) & "-" & _
                            objKtbnStrc.strcSelection.strOpSymbol(5) & _
                            objKtbnStrc.strcSelection.strOpSymbol(6) & _
                            objKtbnStrc.strcSelection.strOpSymbol(7)
                        If objKtbnStrc.strcSelection.strOpSymbol(8) = "" Then
                            strKataban = strKataban & strCalibMix
                        Else
                            strKataban = strKataban & "-" & objKtbnStrc.strcSelection.strOpSymbol(8) & strCalibMix
                        End If
                    Case "4"
                        strKataban = strSeriesKata & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                            objKtbnStrc.strcSelection.strOpSymbol(2) & "-" & _
                            objKtbnStrc.strcSelection.strOpSymbol(3) & _
                            objKtbnStrc.strcSelection.strOpSymbol(4) & "-" & _
                            objKtbnStrc.strcSelection.strOpSymbol(5) & "-" & _
                            objKtbnStrc.strcSelection.strOpSymbol(6) & _
                            objKtbnStrc.strcSelection.strOpSymbol(7) & _
                            objKtbnStrc.strcSelection.strOpSymbol(8)
                        If objKtbnStrc.strcSelection.strOpSymbol(9) = "" Or _
                           (objKtbnStrc.strcSelection.strOpSymbol(5) = "F" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(5) = "G") Then
                            strKataban = strKataban & strCalibMix
                        Else
                            strKataban = strKataban & "-" & objKtbnStrc.strcSelection.strOpSymbol(9) & strCalibMix
                        End If
                    Case "5", "7", "9"
                        strKataban = strSeriesKata & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                            objKtbnStrc.strcSelection.strOpSymbol(2) & "-" & _
                            objKtbnStrc.strcSelection.strOpSymbol(3) & _
                            objKtbnStrc.strcSelection.strOpSymbol(4) & "-" & _
                            objKtbnStrc.strcSelection.strOpSymbol(5) & "-" & _
                            objKtbnStrc.strcSelection.strOpSymbol(6) & _
                            objKtbnStrc.strcSelection.strOpSymbol(7) & _
                            objKtbnStrc.strcSelection.strOpSymbol(8)
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "F", "G"
                                strKataban = strKataban & strCalibMix
                            Case Else
                                strKataban = strKataban & "-" & objKtbnStrc.strcSelection.strOpSymbol(9) & strCalibMix
                        End Select
                    Case "6"
                        strKataban = strSeriesKata & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                     objKtbnStrc.strcSelection.strOpSymbol(2) & "-" & _
                                     objKtbnStrc.strcSelection.strOpSymbol(3) & _
                                     objKtbnStrc.strcSelection.strOpSymbol(4) & "-" & _
                                     objKtbnStrc.strcSelection.strOpSymbol(5) & "-" & _
                                     objKtbnStrc.strcSelection.strOpSymbol(6) & _
                                     objKtbnStrc.strcSelection.strOpSymbol(7) & _
                                     objKtbnStrc.strcSelection.strOpSymbol(8) & "-" & _
                                     objKtbnStrc.strcSelection.strOpSymbol(9) & strCalibMix
                End Select
            End If

            '不要なハイフンを除去
            strUpdKataban(0) = KHKataban.fncHypenCut(strKataban)
            strUpdUse(0) = "1"
            strUpdPosition(0) = ""
            For inti As Integer = 1 To strUpdKataban.Length - 1
                strUpdKataban(inti) = String.Empty
            Next
            For inti As Integer = 1 To strUpdUse.Length - 1
                strUpdUse(inti) = 0
            Next

            ReDim objKtbnStrc.strcSelection.strOpIsoShowFlag(0)

            '画面の入力内容を編集・セットする
            For intI As Integer = Siyou_05.ElType1 To Siyou_05.SpDecomp4
                Select Case strKeyKata.Trim
                    Case "1", "4", "6"
                        If intI = Siyou_05.ElType1 Or intI = Siyou_05.ElType2 Then
                            strUpdKataban(intI) = String.Empty
                            strUpdUse(intI) = "0"
                            strUpdPosition(intI) = arySelectInf(intI - 1)
                        Else
                            strUpdKataban(intI) = strKataValues(intI - 1)
                            strUpdUse(intI) = strUseValues(intI - 1)
                            strUpdPosition(intI) = arySelectInf(intI - 1)
                        End If
                    Case "8"
                        If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "" Then
                            If intI = Siyou_05.ElType1 Or intI = Siyou_05.ElType2 Then
                                strUpdKataban(intI) = String.Empty
                                strUpdUse(intI) = "0"
                                strUpdPosition(intI) = arySelectInf(intI - 1)
                            Else
                                strUpdKataban(intI) = strKataValues(intI - 1)
                                strUpdUse(intI) = strUseValues(intI - 1)
                                strUpdPosition(intI) = arySelectInf(intI - 1)
                            End If
                        Else
                            strUpdKataban(intI) = strKataValues(intI - 1)
                            strUpdUse(intI) = strUseValues(intI - 1)
                            strUpdPosition(intI) = arySelectInf(intI - 1)
                        End If
                    Case Else
                        strUpdKataban(intI) = strKataValues(intI - 1)
                        strUpdUse(intI) = strUseValues(intI - 1)
                        strUpdPosition(intI) = arySelectInf(intI - 1)

                        ReDim Preserve objKtbnStrc.strcSelection.strOpIsoShowFlag(UBound(objKtbnStrc.strcSelection.strOpIsoShowFlag) + 1)
                        objKtbnStrc.strcSelection.strOpIsoShowFlag(UBound(objKtbnStrc.strcSelection.strOpIsoShowFlag)) = "Y"

                End Select
            Next

            ''画面の入力内容を編集・セットする
            '流露遮蔽板(給気)
            strUpdKataban(Siyou_05.ExpCovRep) = strKataValues(Siyou_05.ExpCovRep - 1)
            strUpdUse(Siyou_05.ExpCovRep) = strUseValues(Siyou_05.ExpCovRep - 1)
            strUpdPosition(Siyou_05.ExpCovRep) = arySelectInf(Siyou_05.ExpCovRep - 1)
            '流露遮蔽板(排気)
            strUpdKataban(Siyou_05.ExpCovExh) = strKataValues(Siyou_05.ExpCovExh - 1)
            strUpdUse(Siyou_05.ExpCovExh) = strUseValues(Siyou_05.ExpCovExh - 1) * 2
            strUpdPosition(Siyou_05.ExpCovExh) = arySelectInf(Siyou_05.ExpCovExh - 1)

            With "接続ブロック編集・セット"

                '接続ブロックの使用数の追加
                objKtbnStrc.strcSelection.intQuantity = fncAddMixBlockQuantity(strUpdUse, objKtbnStrc, intCount, sbCoordinates)

                '接続ブロックの形番の追加
                objKtbnStrc.strcSelection.strOptionKataban = fncAddMixBlockKataban(strUpdKataban, strKataValues, arySelectInf, objKtbnStrc)

                '接続ブロックの設置位置の追加
                objKtbnStrc.strcSelection.strPositionInfo = fncGetMixBlockPosition(strUpdPosition, intCount, sbCoordinates)

            End With

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 引当仕様書情報編集(LMFシリーズ)
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intMode"></param>
    ''' <remarks></remarks>
    Public Shared Sub subEditSpecInfoLMF(ByRef objKtbnStrc As KHKtbnStrc, intMode As Integer)
        Dim strKataban As String = ""
        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban

        Try
            Dim strUpdUse(UBound(objKtbnStrc.strcSelection.intQuantity) + 2) As Double
            Dim strUpdKataban(UBound(objKtbnStrc.strcSelection.strOptionKataban) + 2) As String
            'CHANGED BY YGY 20141119    CMFZ2-HX3-HY3BDU  仕様書出力
            'Dim strUpdPosition(UBound(objKtbnStrc.strcSelection.strPositionInfo) + 1) As String
            Dim strUpdPosition(objKtbnStrc.strcSelection.strPositionInfo.Count + 1) As String

            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo


            'CHANGED BY YGY 20141015
            ReDim strUpdUse(UBound(objKtbnStrc.strcSelection.intQuantity) + 1)
            ReDim strUpdKataban(UBound(objKtbnStrc.strcSelection.strOptionKataban) + 1)
            ReDim strUpdPosition(UBound(objKtbnStrc.strcSelection.strPositionInfo) + 1)

            'ペース形番設定
            strKataban = strSeriesKata & _
                         objKtbnStrc.strcSelection.strOpSymbol(1).ToString & "-" & _
                         objKtbnStrc.strcSelection.strOpSymbol(2).ToString & "-" & _
                         objKtbnStrc.strcSelection.strOpSymbol(3).ToString & "-" & _
                         objKtbnStrc.strcSelection.strOpSymbol(4).ToString

            '不要なハイフンを除去
            strUpdKataban(0) = KHKataban.fncHypenCut(strKataban)
            strUpdUse(0) = "1"
            strUpdPosition(0) = ""

            'ABポート接続口径設定
            For intI As Integer = 0 To 9
                For intJ As Integer = Siyou_06.Elect1 To Siyou_06.Elect6
                    If arySelectInf(intJ - 1)(intI) = "1" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            Case "01"
                                arySelectInf(Siyou_06.ABCon01 - 1) = StrNewString(arySelectInf(Siyou_06.ABCon01 - 1), intI, "1")
                                strUseValues(Siyou_06.ABCon01 - 1) = CInt(strUseValues(Siyou_06.ABCon01 - 1)) + 1
                            Case "02"
                                arySelectInf(Siyou_06.ABCon02 - 1) = StrNewString(arySelectInf(Siyou_06.ABCon02 - 1), intI, "1")
                                strUseValues(Siyou_06.ABCon02 - 1) = CInt(strUseValues(Siyou_06.ABCon02 - 1)) + 1
                            Case "C4"
                                arySelectInf(Siyou_06.ABCon04 - 1) = StrNewString(arySelectInf(Siyou_06.ABCon04 - 1), intI, "1")
                                strUseValues(Siyou_06.ABCon04 - 1) = CInt(strUseValues(Siyou_06.ABCon04 - 1)) + 1
                            Case "C6"
                                arySelectInf(Siyou_06.ABCon06 - 1) = StrNewString(arySelectInf(Siyou_06.ABCon06 - 1), intI, "1")
                                strUseValues(Siyou_06.ABCon06 - 1) = CInt(strUseValues(Siyou_06.ABCon06 - 1)) + 1
                            Case "01Z"
                                arySelectInf(Siyou_06.ABCon1Z - 1) = StrNewString(arySelectInf(Siyou_06.ABCon1Z - 1), intI, "1")
                                strUseValues(Siyou_06.ABCon1Z - 1) = CInt(strUseValues(Siyou_06.ABCon1Z - 1)) + 1
                            Case "XX"
                        End Select
                    End If
                Next
            Next

            '画面の入力内容を編集・セットする
            For intI As Integer = Siyou_06.Elect1 To Siyou_06.Pilot
                strUpdKataban(intI) = strKataValues(intI - 1)
                strUpdUse(intI) = strUseValues(intI - 1)
                strUpdPosition(intI) = arySelectInf(intI - 1)
            Next
            '流露遮蔽板(給気)
            strUpdKataban(Siyou_06.ExpCovRep) = strKataValues(Siyou_06.ExpCovRep - 1)
            strUpdUse(Siyou_06.ExpCovRep) = strUseValues(Siyou_06.ExpCovRep - 1)
            strUpdPosition(Siyou_06.ExpCovRep) = arySelectInf(Siyou_06.ExpCovRep - 1)
            '流露遮蔽板(排気)
            strUpdKataban(Siyou_06.ExpCovExh) = strKataValues(Siyou_06.ExpCovExh - 1)
            strUpdUse(Siyou_06.ExpCovExh) = strUseValues(Siyou_06.ExpCovExh - 1) * 2
            strUpdPosition(Siyou_06.ExpCovExh) = arySelectInf(Siyou_06.ExpCovExh - 1)

            '情報追加
            objKtbnStrc.strcSelection.intQuantity = strUpdUse
            objKtbnStrc.strcSelection.strOptionKataban = strUpdKataban
            objKtbnStrc.strcSelection.strPositionInfo = strUpdPosition

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 接続ブロック形番の作成
    ''' </summary>
    ''' <param name="strUpdKataban">追加前の形番</param>
    ''' <param name="strKataValues"></param>
    ''' <param name="arySelectInf"></param>
    ''' <param name="objKtbnStrc"></param>
    ''' <returns>接続ブロックの形番を追加した形番</returns>
    ''' <remarks></remarks>
    Private Shared Function fncAddMixBlockKataban(ByVal strUpdKataban As String(), _
                                                  ByVal strKataValues As String(), _
                                                  ByVal arySelectInf As String(), _
                                                  ByVal objKtbnStrc As KHKtbnStrc) As String()

        Dim strResult As String() = strUpdKataban
        Dim strSeries As String = String.Empty

        If objKtbnStrc.strcSelection.strSeriesKataban.StartsWith("GMF") Then
            strSeries = "GMF"
        Else
            strSeries = "CMF"
        End If

        'GMFシリーズ
        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "Z" Then
            For intI As Integer = 0 To 9
                If arySelectInf(Siyou_05.ElType1 - 1)(intI) = "1" Then
                    If Left(strKataValues(Siyou_05.ElType1 - 1).Trim, 5) = "PV5-6" Or _
                       Left(strKataValues(Siyou_05.ElType1 - 1).Trim, 6) = "PV5G-6" Or _
                       Left(strKataValues(Siyou_05.ElType1 - 1).Trim, 3) = "CM1" Then
                        strResult(UBound(strResult) - 1) = strSeries & "BZ-00L"
                    Else
                        strResult(UBound(strResult) - 1) = strSeries & "BZ-00R"
                    End If
                    Exit For
                End If
                If arySelectInf(Siyou_05.ElType2 - 1)(intI) = "1" Then
                    If Left(strKataValues(Siyou_05.ElType2 - 1).Trim, 5) = "PV5-6" Or _
                       Left(strKataValues(Siyou_05.ElType2 - 1).Trim, 6) = "PV5G-6" Or _
                       Left(strKataValues(Siyou_05.ElType1 - 1).Trim, 3) = "CM1" Then
                        strResult(UBound(strResult) - 1) = strSeries & "BZ-00L"
                    Else
                        strResult(UBound(strResult) - 1) = strSeries & "BZ-00R"
                    End If
                    Exit For
                End If
                If arySelectInf(Siyou_05.ElType3 - 1)(intI) = "1" Then
                    If Left(strKataValues(Siyou_05.ElType3 - 1).Trim, 5) = "PV5-6" Or _
                       Left(strKataValues(Siyou_05.ElType3 - 1).Trim, 6) = "PV5G-6" Or _
                       Left(strKataValues(Siyou_05.ElType1 - 1).Trim, 3) = "CM1" Then
                        strResult(UBound(strResult) - 1) = strSeries & "BZ-00L"
                    Else
                        strResult(UBound(strResult) - 1) = strSeries & "BZ-00R"
                    End If
                    Exit For
                End If
                If arySelectInf(Siyou_05.ElType4 - 1)(intI) = "1" Then
                    If Left(strKataValues(Siyou_05.ElType4 - 1).Trim, 5) = "PV5-6" Or _
                       Left(strKataValues(Siyou_05.ElType4 - 1).Trim, 6) = "PV5G-6" Or _
                       Left(strKataValues(Siyou_05.ElType1 - 1).Trim, 3) = "CM1" Then
                        strResult(UBound(strResult) - 1) = strSeries & "BZ-00L"
                    Else
                        strResult(UBound(strResult) - 1) = strSeries & "BZ-00R"
                    End If
                    Exit For
                End If
                If arySelectInf(Siyou_05.ElType5 - 1)(intI) = "1" Then
                    If Left(strKataValues(Siyou_05.ElType5 - 1).Trim, 5) = "PV5-6" Or _
                       Left(strKataValues(Siyou_05.ElType5 - 1).Trim, 6) = "PV5G-6" Or _
                       Left(strKataValues(Siyou_05.ElType1 - 1).Trim, 3) = "CM1" Then
                        strResult(UBound(strResult) - 1) = strSeries & "BZ-00L"
                    Else
                        strResult(UBound(strResult) - 1) = strSeries & "BZ-00R"
                    End If
                    Exit For
                End If
                If arySelectInf(Siyou_05.ElType6 - 1)(intI) = "1" Then
                    If Left(strKataValues(Siyou_05.ElType6 - 1).Trim, 5) = "PV5-6" Or _
                       Left(strKataValues(Siyou_05.ElType6 - 1).Trim, 6) = "PV5G-6" Or _
                       Left(strKataValues(Siyou_05.ElType1 - 1).Trim, 3) = "CM1" Then
                        strResult(UBound(strResult) - 1) = strSeries & "BZ-00L"
                    Else
                        strResult(UBound(strResult) - 1) = strSeries & "BZ-00R"
                    End If
                    Exit For
                End If
            Next
        Else
            strResult(UBound(strResult) - 1) = ""
        End If

        Return strResult

    End Function

    ''' <summary>
    ''' 接続ブロック設置位置の作成
    ''' </summary>
    ''' <param name="strUpdPosition">追加前の設置位置</param>
    ''' <param name="intCount"></param>
    ''' <param name="sbCoordinates"></param>
    ''' <returns>接続ブロックの設置位置を追加した設置位置</returns>
    ''' <remarks></remarks>
    Private Shared Function fncGetMixBlockPosition(ByVal strUpdPosition As String(), _
                                                   ByVal intCount As Integer, _
                                                   ByVal sbCoordinates As StringBuilder) As String()
        Dim strResult As String() = strUpdPosition

        If intCount > 0 Then

            Dim strPosCoord As String()
            Dim strPosInfo As String = "0000000000"

            '接続ブロックの座標を取得
            strPosCoord = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1).Split(strPipe)

            '列Noを設置位置情報に反映
            For intI As Integer = 0 To UBound(strPosCoord)

                Dim strWork As String() = strPosCoord(intI).Split(strComma)
                strPosInfo = StrNewString(strPosInfo, Int(strWork(1)) - 2, "1")

            Next

            strResult(UBound(strResult) - 1) = strPosInfo

        Else

            strResult(UBound(strResult) - 1) = String.Empty

        End If

        Return strResult

    End Function

    ''' <summary>
    ''' 接続ブロックの使用数の作成
    ''' </summary>
    ''' <param name="dblUpdUse">追加前の仕様数</param>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intCount"></param>
    ''' <param name="sbCoordinates"></param>
    ''' <returns>接続ブロックの使用数を追加した使用数</returns>
    ''' <remarks></remarks>
    Private Shared Function fncAddMixBlockQuantity(ByVal dblUpdUse As Double(), _
                                                   ByVal objKtbnStrc As KHKtbnStrc, _
                                                   ByRef intCount As Integer, _
                                                   ByRef sbCoordinates As StringBuilder) As Double()

        Dim dblResult As Double() = dblUpdUse

        intCount = ClsInputCheck_05.fncGetConectBlockCnt(objKtbnStrc, sbCoordinates, CInt(objKtbnStrc.strcSelection.strOpSymbol(2)))

        dblResult(UBound(dblResult) - 1) = intCount

        Return dblResult

    End Function

    ''' <summary>
    ''' CX継手の区分を取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function fncGetCXKbn(ByVal strSeries As String, _
                                 ByVal strKeyKataban As String) As Integer

        Dim intResult As Integer = 0

        If (strSeries = "M3GA1" And strKeyKataban = "S") Or _
            (strSeries = "M3GA1" And strKeyKataban = "V") Or _
            (strSeries = "M3GB1" And strKeyKataban = "S") Or _
            (strSeries = "M3GB1" And strKeyKataban = "V") Or _
            (strSeries = "M4GA1" And strKeyKataban = "S") Or _
            (strSeries = "M4GA1" And strKeyKataban = "V") Or _
            (strSeries = "M4GB1" And strKeyKataban = "S") Or _
            (strSeries = "M4GB1" And strKeyKataban = "V") Or _
            (strSeries = "MN3GA1" And strKeyKataban = "S") Or _
            (strSeries = "MN3GA1" And strKeyKataban = "V") Or _
            (strSeries = "MN3GB1" And strKeyKataban = "S") Or _
            (strSeries = "MN3GB1" And strKeyKataban = "V") Or _
            (strSeries = "MN4GA1" And strKeyKataban = "S") Or _
            (strSeries = "MN4GA1" And strKeyKataban = "V") Or _
            (strSeries = "MN4GB1" And strKeyKataban = "S") Or _
            (strSeries = "MN4GB1" And strKeyKataban = "V") Then

            intResult = 1

        ElseIf (strSeries = "M3GA2" And strKeyKataban = "S") Or _
            (strSeries = "M3GA2" And strKeyKataban = "V") Or _
            (strSeries = "M3GB2" And strKeyKataban = "S") Or _
            (strSeries = "M3GB2" And strKeyKataban = "V") Or _
            (strSeries = "M4GA2" And strKeyKataban = "S") Or _
            (strSeries = "M4GA2" And strKeyKataban = "V") Or _
            (strSeries = "M4GB2" And strKeyKataban = "S") Or _
            (strSeries = "M4GB2" And strKeyKataban = "V") Or _
            (strSeries = "MN3GA2" And strKeyKataban = "S") Or _
            (strSeries = "MN3GA2" And strKeyKataban = "V") Or _
            (strSeries = "MN3GB2" And strKeyKataban = "S") Or _
            (strSeries = "MN3GB2" And strKeyKataban = "V") Or _
            (strSeries = "MN4GA2" And strKeyKataban = "S") Or _
            (strSeries = "MN4GA2" And strKeyKataban = "V") Or _
            (strSeries = "MN4GB2" And strKeyKataban = "S") Or _
            (strSeries = "MN4GB2" And strKeyKataban = "V") Then

            intResult = 2

        ElseIf (strSeries = "M3GA3" And strKeyKataban = "S") Or _
            (strSeries = "M3GA3" And strKeyKataban = "V") Or _
            (strSeries = "M3GB3" And strKeyKataban = "S") Or _
            (strSeries = "M3GB3" And strKeyKataban = "V") Or _
            (strSeries = "M4GA3" And strKeyKataban = "S") Or _
            (strSeries = "M4GA3" And strKeyKataban = "V") Or _
            (strSeries = "M4GB3" And strKeyKataban = "S") Or _
            (strSeries = "M4GB3" And strKeyKataban = "V") Then

            intResult = 3

        End If

        Return intResult

    End Function

End Class
