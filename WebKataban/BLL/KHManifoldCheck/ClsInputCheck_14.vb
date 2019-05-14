Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_14

    Public Shared intPosRowCnt As Integer = 6
    Public Shared intColCnt As Integer = 25

    '********************************************************************************************
    '*【関数名】
    '*   fncInpCheck
    '*【処理】
    '*   入力チェック
    '*【引数】
    '*   strKataValues  : 形番の選択値配列          strUseValues    : 使用数の入力値配列
    '*   arySelectInf   : 設置位置の選択値配列
    '*【更新】
    '********************************************************************************************
    Public Shared Function fncInputChk(objKtbnStrc As KHKtbnStrc, ByRef dblStdNum As Double, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean

        Dim sbCoordinates As New System.Text.StringBuilder
        Dim intColR As Integer = 0
        Dim bolChkFlag As Boolean = False
        Dim intExhaustCnt As Integer = 0            '電装・給排気ブロック使用数計
        Dim bolEndRightChk As Boolean = False       'エンドブロック右列チェック
        Dim intSolenoidCnt As Integer = 0           'ソレノイドカウント
        Dim intEvtCnt As Integer = 0                'EVT連続カウント
        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban

        fncInputChk = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            '------ 既存システムで、コントロール上の値変更時にチェックしていた内容 ----------
            '8.1 形番選択チェック
            For intRI As Integer = 1 To intPosRowCnt
                If Trim(strKataValues(intRI - 1)).Length = 0 Then
                    For intCI As Integer = 1 To intColCnt
                        If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                            sbCoordinates.Append(CStr(intRI) & strComma & "0")
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W1400"
                            Exit Function
                        End If
                    Next
                End If
            Next

            '8.2 数値テキスト入力値チェック
            If Not SiyouBLL.fncOtherKataCheck(objKtbnStrc, Siyou_14.Plug1, Siyou_14.Silencer, 0, strMsgCd) Then
                Exit Function
            End If

            '8.3 取付レール長さ設定値チェック
            If strKataValues(Siyou_14.Rail - 1).ToString.Length <= 0 Then strKataValues(Siyou_14.Rail - 1) = 0
            If Not SiyouBLL.fncRailchk(strKataValues(Siyou_14.Rail - 1), strUseValues(Siyou_14.Rail - 1), dblStdNum, strMsgCd) Then
                strMsg = Siyou_14.Rail & ",0"
                Exit Function
            End If

            '----- 入力内容チェック ----------
            '最右部列取得
            For intCI As Integer = intColCnt To 1 Step -1
                For intRI As Integer = 1 To intPosRowCnt
                    If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                        intColR = intCI
                        Exit For
                    End If
                Next
                If intColR > 0 Then
                    Exit For
                End If
            Next

            '7.1 接続位置チェック
            For intCI As Integer = 1 To intColR
                For intRI As Integer = 1 To intPosRowCnt
                    If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                        bolChkFlag = True
                    End If
                Next
                If bolChkFlag = False Then
                    sbCoordinates.Append("0" & strComma & CStr(intCI))
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1020"
                    Exit Function
                End If
                bolChkFlag = False
            Next

            '電装・給排気ブロック使用数計
            For intRI As Integer = Siyou_14.Exhaust1 To Siyou_14.Exhaust3
                intExhaustCnt = intExhaustCnt + CInt(strUseValues(intRI - 1))
            Next
            Select Case intExhaustCnt
                Case 0
                    '7.2 電装・給排気ブロック必須チェック
                    sbCoordinates.Append(CStr(Siyou_14.Exhaust1) & strComma & "0")
                    If Trim(strKataValues(Siyou_14.Exhaust2 - 1)).Length > 0 Then
                        sbCoordinates.Append(strPipe & CStr(Siyou_14.Exhaust2) & strComma & "0")
                    End If
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W2370"
                    Exit Function
                Case Is >= 4
                    '7.3 電装・給排気ブロック複数指定チェック
                    If CInt(strUseValues(Siyou_14.Exhaust1 - 1)) > 0 Then
                        sbCoordinates.Append(CStr(Siyou_14.Exhaust1) & strComma & "0")
                    End If
                    If CInt(strUseValues(Siyou_14.Exhaust2 - 1)) > 0 Then
                        If sbCoordinates.Length > 0 Then
                            sbCoordinates.Append(strPipe)
                        End If
                        sbCoordinates.Append(CStr(Siyou_14.Exhaust2) & strComma & "0")
                    End If
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W2380"
                    Exit Function
            End Select

            '7.4 左側エンドブロック複数指定チェック
            If CInt(strUseValues(Siyou_14.End1 - 1)) > 1 Then
                sbCoordinates.Append(CStr(Siyou_14.End1) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W1100"
                Exit Function
            End If

            '7.5 右側エンドブロック必須チェック
            Select Case objKtbnStrc.strcSelection.strOpSymbol(4)
                Case "T11R", "T30R"
                    '最右選択位置が2～4行目以外の場合、エラー
                    For intRI As Integer = Siyou_14.Exhaust1 To Siyou_14.Exhaust3
                        If arySelectInf(intRI - 1)(intColR - 1) = "1" Then
                            bolEndRightChk = True
                            Exit For
                        End If
                    Next
                    If bolEndRightChk = False Then
                        If intColR < intColCnt Then
                            sbCoordinates.Append(CStr(Siyou_14.Exhaust1) & strComma & CStr(intColR + 1))
                            If Trim(strKataValues(Siyou_14.Exhaust2 - 1)).Length > 0 Then
                                sbCoordinates.Append(strPipe & CStr(Siyou_14.Exhaust2) & strComma & CStr(intColR + 1))
                            End If
                        End If
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W2390"
                        Exit Function
                    End If
                Case "T9DAR", "T9GAR", "T9L0R"
                    '最右選択位置が6行目以外の場合、エラー
                    If arySelectInf(Siyou_14.End2 - 1)(intColR - 1) = "1" Then
                    Else
                        If intColR < intColCnt Then
                            sbCoordinates.Append(CStr(Siyou_14.End2) & strComma & CStr(intColR + 1))
                        End If
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W2400"
                        Exit Function
                    End If
            End Select

            '7.6 右側エンドブロック複数指定チェック
            If CInt(strUseValues(Siyou_14.End2 - 1)) > 1 Then
                sbCoordinates.Append(CStr(Siyou_14.End2) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W1100"
                Exit Function
            End If

            '7.7 右側エンドブロック指定時右端設置チェック
            If CInt(strUseValues(Siyou_14.End2 - 1)) > 0 Then
                If arySelectInf(Siyou_14.End2 - 1)(intColR - 1) = "1" Then
                Else
                    sbCoordinates.Append(CStr(Siyou_14.End2) & strComma & CStr(intColR))
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1650"
                    Exit Function
                End If
            End If

            '7.8 EVT連数チェック
            If CInt(strUseValues(Siyou_14.Evt - 1)) > CInt(objKtbnStrc.strcSelection.strOpSymbol(5).ToString) Then
                strMsgCd = "W2410"
                Exit Function
            End If
            If CInt(strUseValues(Siyou_14.Evt - 1)) < CInt(objKtbnStrc.strcSelection.strOpSymbol(5).ToString) Then
                strMsgCd = "W2420"
                Exit Function
            End If

            '7.9 電装・給排気ブロック：EVT組合せチェック
            If CInt(strUseValues(Siyou_14.Evt - 1)) > (8 * intExhaustCnt) Then
                strMsgCd = "W2430"
                Exit Function
            End If

            '7.10 シリアル電装NET点数チェック
            If Left(objKtbnStrc.strcSelection.strOpSymbol(4), 2) = "T9" Then
                For intRI As Integer = Siyou_14.Exhaust1 To Siyou_14.Exhaust3
                    If CInt(strUseValues(intRI - 1)) > 0 Then
                        Select Case strKataValues(intRI - 1).Substring(7, 1)
                            Case "A"
                                intSolenoidCnt = intSolenoidCnt + 4 * CInt(strUseValues(intRI - 1))
                            Case "0"
                                intSolenoidCnt = intSolenoidCnt + 8 * CInt(strUseValues(intRI - 1))
                        End Select
                    End If
                Next
                If intSolenoidCnt < CInt(objKtbnStrc.strcSelection.strOpSymbol(5).ToString) Then
                    strMsgCd = "W2440"
                    Exit Function
                End If
            End If

            '7.11 EVT連続設置チェック
            'EVTが9個以上ある場合のみチェックを行う
            If CInt(strUseValues(Siyou_14.Evt - 1)) > 8 Then
                For intCI As Integer = 1 To intColR
                    For intRI As Integer = Siyou_14.Evt To Siyou_14.End1
                        If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                            If intRI = Siyou_14.Evt Then
                                'EVT連続選択数をインクリメント
                                intEvtCnt = intEvtCnt + 1
                            Else
                                'EVT連続選択数リセット
                                intEvtCnt = 0
                            End If
                            Exit For
                        End If
                    Next
                    'EVT選択数が9個以上連続の場合、エラー
                    If intEvtCnt = 9 Then
                        strMsgCd = "W2450"
                        Exit Function
                    End If
                Next
            End If

            fncInputChk = True

        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try
    End Function
End Class
