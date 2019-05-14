Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_10

    Public Shared intPosRowCnt As Integer = 14
    Public Shared intColCnt As Integer = 25

    Public Shared Function fncInputChk(objKtbnStrc As KHKtbnStrc, ByRef dblStdNum As Double, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        fncInputChk = False
        Try
            '入力チェック
            If Not fncInpCheck2(objKtbnStrc, dblStdNum, strMsg, strMsgCd) Then
                '画面作成
                Exit Function
            End If

            '入力チェック2
            If Not fncInpCheck1(objKtbnStrc, dblStdNum, strMsg, strMsgCd) Then
                '画面作成
                Exit Function
            End If
            fncInputChk = True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncInputChk_1
    '*【処理】
    '*   入力チェック①
    '********************************************************************************************
    Public Shared Function fncInpCheck1(objKtbnStrc As KHKtbnStrc, ByRef dblStdNum As Double, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        Dim bolFlg1 As Boolean
        Dim intRightCol As Integer
        Dim bolRow_9 As Boolean
        Dim bolRow_10 As Boolean
        Dim bolRow_13 As Boolean
        Dim bolRow_14 As Boolean
        Dim intSoreCnt As Integer
        Dim intSelCnt As Integer
        Dim intMixPosDv(11) As Integer
        Dim strKtbn As String
        Dim intInpUse As Integer
        Dim intMixCnt As Integer
        Dim intExhltBlkCnt As Integer
        Dim bolRtn As Boolean
        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban

        fncInpCheck1 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            '--------------------------------------------------------------
            '接続位置ﾁｪｯｸ
            '-- 選択されているｾﾙのうち最右列を取得する
            For intC As Integer = intColCnt - 1 To 0 Step -1
                bolFlg1 = False
                For intR As Integer = 0 To intPosRowCnt - 1
                    If arySelectInf(intR)(intC) = "1" Then
                        bolFlg1 = True
                        Exit For
                    End If
                Next
                If bolFlg1 Then
                    intRightCol = intC
                    Exit For
                End If
            Next

            For intC As Integer = 0 To intRightCol
                bolFlg1 = False
                For intR As Integer = 0 To intPosRowCnt - 1
                    If arySelectInf(intR)(intC) = "1" Then
                        bolFlg1 = True
                        Exit For
                    End If
                Next
                If Not bolFlg1 Then
                    'message:W1020=選択されない接続位置があります。接続位置=[1]
                    strMsg = "0" & strComma & CStr(intC + 1)
                    strMsgCd = "W1020"
                    Exit Try
                End If
            Next

            '--------------------------------------------------------------
            '左側ｴﾝﾄﾞﾌﾞﾛｯｸ必須ﾁｪｯｸ
            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).ToString
                Case "T10R", "T11R", "T30R", "T50R", "C", "C0", "C1", "C2"
                    Select Case strKataValues(12)
                        Case "N4S0-EL", "N4S0-EXL"
                        Case Else
                            'message:W2640=左側取付エンドブロック（N4S0-ELまたはN4S0-EXL）の形番を選択して下さい。
                            strMsgCd = "W2640"
                            Exit Try
                    End Select
            End Select

            '--------------------------------------------------------------
            '右端必須ﾁｪｯｸ
            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).ToString
                Case "T10", "T11", "T30", "T50", "T621", "T631", "T6A0", "T6A1", "T6C0", "T6C1", "T6E0", "T6E1", "T6G1", "T6J0", "T6J1", "T6K1"

                    If Not arySelectInf(12)(intRightCol) = "1" Then
                        'message:W2650=右側取付エンドブロック（N4S0-EまたはN4S0-EX）を最後に指定して下さい。接続位置=[1]
                        strMsg = "13,0"
                        strMsgCd = "W2650"
                        Exit Try
                    End If

                Case "C", "C0", "C1", "C2"

                    If Not arySelectInf(13)(intRightCol) = "1" Then
                        'message:W2650=右側取付エンドブロック（N4S0-EまたはN4S0-EX）を最後に指定して下さい。接続位置=[1]
                        strMsg = "14,0"
                        strMsgCd = "W2650"
                        Exit Try
                    End If

                Case "T10R", "T11R", "T30R", "T50R"

                    If Not arySelectInf(0)(intRightCol) = "1" Then
                        'message:W2660=右側取付配線ブロックを最後に指定して下さい。接続位置=[1]
                        strMsg = "1,0"
                        strMsgCd = "W2660"
                        Exit Try
                    End If

            End Select

            '--------------------------------------------------------------
            '給排気ﾌﾞﾛｯｸ・ｴﾝﾄﾞﾌﾞﾛｯｸ組合せﾁｪｯｸ
            '-- 9行目ﾁｪｯｸ
            bolRow_9 = False
            If CInt(strUseValues(8)) > 0 Then
                If strKataValues(8).IndexOf("X") > -1 Then
                    bolRow_9 = True
                End If
            End If
            '-- 10行目ﾁｪｯｸ
            bolRow_10 = False
            If CInt(strUseValues(9)) > 0 Then
                If strKataValues(9).IndexOf("X") > -1 Then
                    bolRow_10 = True
                End If
            End If

            If bolRow_9 = True Or bolRow_10 = True Then
                '-- 13行目ﾁｪｯｸ
                bolRow_13 = False
                If CInt(strUseValues(12)) > 0 Then
                    Select Case strKataValues(12)
                        Case "N4S0-EX", "N4S0-EXL"
                            bolRow_13 = True
                    End Select
                End If
                '-- 14行目ﾁｪｯｸ
                bolRow_14 = False
                If CInt(strUseValues(13)) > 0 Then
                    Select Case strKataValues(13)
                        Case "N4S0-EX", "N4S0-EXL"
                            bolRow_14 = True
                    End Select
                End If

                If bolRow_13 = False And bolRow_14 = False Then
                    'message:W2670=エンドブロック（N4S0-EXLまたはN4S0-EX）を指定して下さい。
                    strMsg = "13,0"
                    strMsgCd = "W2670"
                    Exit Try
                End If
            End If

            '--------------------------------------------------------------
            '配線ﾌﾞﾛｯｸ複数指定ﾁｪｯｸ
            If strUseValues(0) > 1 Then
                'message:W1530=配線ブロックを複数指定することはできません。
                strMsg = "1,0"
                strMsgCd = "W1530"
                Exit Try
            End If

            '--------------------------------------------------------------
            'ｿﾚﾉｲﾄﾞ点数ﾁｪｯｸ
            Select Case strSeriesKata
                Case "MN3S0", "MN4S0"
                    'ﾊﾞﾙﾌﾞﾌﾞﾛｯｸが選択されているか(2～8行目)
                    For intR As Integer = 1 To 7
                        If CInt(strUseValues(intR)) > 0 Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).ToString
                                Case "C", "C0", "C1", "C2"
                                Case Else
                                    If Mid(strKataValues(intR), 5, 1) = "1" Then
                                        intSoreCnt = intSoreCnt + (1 * CInt(strUseValues(intR)))
                                    Else
                                        If Not Mid(strKataValues(intR), 5, 1) = "-" Then
                                            intSoreCnt = intSoreCnt + (2 * CInt(strUseValues(intR)))
                                        End If
                                    End If
                            End Select
                        End If
                    Next

                    If intSoreCnt > KHKataban.fncGetMaxSol(objKtbnStrc.strcSelection.strOpSymbol, 10) Then
                        'message:W1150=ソレノイド点数が多すぎます。
                        strMsgCd = "W1150"
                        Exit Try
                    End If

            End Select

            '--------------------------------------------------------------
            'ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ連数ﾁｪｯｸ
            For idx As Integer = 0 To 11
                'ﾐｯｸｽﾁｪｯｸ値(切替位置区分)
                intMixPosDv(idx) = 0
            Next
            '電磁弁連数値
            intSelCnt = 0

            For intR As Integer = 1 To 7
                strKtbn = strKataValues(intR)
                intInpUse = CInt(strUseValues(intR))
                If (Not strKtbn = String.Empty) And intInpUse >= 1 Then

                    '-- 切替位置区分1文字目 = "8"
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(1).ToString, 1) = "8" Then

                        If Mid(strKtbn, 2, 1) = "4" And Mid(strKtbn, 5, 2) = "10" Then
                            intMixPosDv(1) = intMixPosDv(1) + intInpUse
                        End If
                        If Mid(strKtbn, 2, 1) = "4" And Mid(strKtbn, 5, 2) = "20" Then
                            intMixPosDv(2) = intMixPosDv(2) + intInpUse
                        End If
                        If Mid(strKtbn, 2, 1) = "4" And Mid(strKtbn, 5, 2) = "30" Then
                            intMixPosDv(3) = intMixPosDv(3) + intInpUse
                        End If
                        If Mid(strKtbn, 2, 1) = "4" And Mid(strKtbn, 5, 2) = "40" Then
                            intMixPosDv(4) = intMixPosDv(4) + intInpUse
                        End If
                        If Mid(strKtbn, 2, 1) = "4" And Mid(strKtbn, 5, 2) = "50" Then
                            intMixPosDv(5) = intMixPosDv(5) + intInpUse
                        End If

                        If Mid(strKtbn, 2, 1) = "3" And Mid(strKtbn, 5, 2) = "10" Then
                            intMixPosDv(6) = intMixPosDv(6) + intInpUse
                        End If
                        If Mid(strKtbn, 2, 1) = "3" And Mid(strKtbn, 5, 2) = "11" Then
                            intMixPosDv(7) = intMixPosDv(7) + intInpUse
                        End If
                        If Mid(strKtbn, 2, 1) = "3" And Mid(strKtbn, 5, 2) = "66" Then
                            intMixPosDv(8) = intMixPosDv(8) + intInpUse
                        End If
                        If Mid(strKtbn, 2, 1) = "3" And Mid(strKtbn, 5, 2) = "67" Then
                            intMixPosDv(9) = intMixPosDv(9) + intInpUse
                        End If
                        If Mid(strKtbn, 2, 1) = "3" And Mid(strKtbn, 5, 2) = "76" Then
                            intMixPosDv(10) = intMixPosDv(10) + intInpUse
                        End If
                        If Mid(strKtbn, 2, 1) = "3" And Mid(strKtbn, 5, 2) = "77" Then
                            intMixPosDv(11) = intMixPosDv(11) + intInpUse
                        End If

                    End If

                    '電磁弁連数値に対象行の選択数を追加
                    intSelCnt = intSelCnt + intInpUse

                End If
            Next

            '電磁弁Max値と比較
            If intSelCnt > CInt(objKtbnStrc.strcSelection.strOpSymbol(7).ToString) Then
                'message:W1170=選択した電磁弁の連数が指定した値より多いです。
                strMsgCd = "W1170"
                Exit Try
            End If
            If intSelCnt < CInt(objKtbnStrc.strcSelection.strOpSymbol(7).ToString) Then
                'message:W1180=選択した電磁弁の連数が指定した値より少ないです。
                strMsgCd = "W1180"
                Exit Try
            End If

            intMixCnt = 0
            If Left(objKtbnStrc.strcSelection.strOpSymbol(1).ToString, 1) = "8" Then
                For idx As Integer = 1 To 11
                    If intMixPosDv(idx) >= 1 Then
                        intMixCnt = intMixCnt + 1
                    End If
                Next
                If intMixCnt <= 1 Then
                    'message:W1190=電磁弁の切換位置は２種類以上選択してください。
                    strMsgCd = "W1190"
                    Exit Try
                End If
            End If
            intMixCnt = 0
            If Mid(objKtbnStrc.strcSelection.strOpSymbol(3).ToString, 2, 1) = "X" Then
                If Not SiyouBLL.fncMixBlockCheck(objKtbnStrc, 1, 7, strMsgCd) Then
                    Exit Function
                End If
            End If

            '--------------------------------------------------------------
            'ｴﾝﾄﾞﾌﾞﾛｯｸ複数指定ﾁｪｯｸ
            If CInt(strUseValues(12)) > 1 Then
                'message:W1100=エンドブロックを複数指定することはできません。
                strMsg = "13,0"
                strMsgCd = "W1100"
                Exit Try
            End If
            If CInt(strUseValues(13)) > 1 Then
                'message:W1100=エンドブロックを複数指定することはできません。
                strMsg = "14,0"
                strMsgCd = "W1100"
                Exit Try
            End If

            '--------------------------------------------------------------
            '給排気ﾌﾞﾛｯｸ 外部ﾊﾟｲﾛｯﾄ選択ﾁｪｯｸ
            intExhltBlkCnt = 0
            Select Case strSeriesKata
                Case "MT3S0", "MT4S0"
                    If CInt(strUseValues(8)) > 0 Then
                        If strKataValues(8).IndexOf("-Q-") > -1 _
                        Or strKataValues(8).IndexOf("-QK-") > -1 _
                        Or strKataValues(8).IndexOf("-QKZ-") > -1 Then
                            intExhltBlkCnt = intExhltBlkCnt + CInt(strUseValues(8))
                        End If
                    End If
                    If CInt(strUseValues(9)) > 0 Then
                        If strKataValues(9).IndexOf("-Q-") > -1 _
                        Or strKataValues(9).IndexOf("-QK-") > -1 _
                        Or strKataValues(9).IndexOf("-QKZ-") > -1 Then
                            intExhltBlkCnt = intExhltBlkCnt + CInt(strUseValues(9))
                        End If
                    End If
                    If intExhltBlkCnt = 0 Then
                        'message:W2680=給排気ブロックに（QまたはQKまたはQKZ）を１ヶ指定して下さい。
                        strMsg = "9,0|10,0"
                        strMsgCd = "W2680"
                        Exit Try
                    End If
                    If intExhltBlkCnt > 1 Then
                        'message:W2690=給排気ブロックに（QまたはQKまたはQKZ）は複数指定できません。
                        strMsg = "9,0|10,0"
                        strMsgCd = "W2690"
                        Exit Try
                    End If
            End Select

            '--------------------------------------------------------------
            '仕切ﾌﾞﾛｯｸ無指定ﾁｪｯｸ
            intExhltBlkCnt = 0
            Select Case strSeriesKata
                Case "MN3S0", "MN4S0"
                    If CInt(strUseValues(10)) = 0 And CInt(strUseValues(11)) = 0 Then
                        If CInt(strUseValues(8)) > 0 Then
                            If strKataValues(8).IndexOf("-Q-") > -1 _
                            Or strKataValues(8).IndexOf("-QK-") > -1 _
                            Or strKataValues(8).IndexOf("-QKZ-") > -1 Then
                                intExhltBlkCnt = intExhltBlkCnt + CInt(strUseValues(8))
                            End If
                        End If
                        If CInt(strUseValues(9)) > 0 Then
                            If strKataValues(9).IndexOf("-Q-") > -1 _
                            Or strKataValues(9).IndexOf("-QK-") > -1 _
                            Or strKataValues(9).IndexOf("-QKZ-") > -1 Then
                                intExhltBlkCnt = intExhltBlkCnt + CInt(strUseValues(9))
                            End If
                        End If
                        If intExhltBlkCnt = 0 Then
                            'message:W2700=給排気ブロック（QまたはQKまたはQKZ）を１ヶ以上指定して下さい。
                            strMsg = "9,0|10,0"
                            strMsgCd = "W2700"
                            Exit Try
                        End If
                    End If
            End Select

            Select Case strSeriesKata
                Case "MN3S0", "MN4S0"
                    '--------------------------------------------------------------
                    '仕切ﾌﾞﾛｯｸ・ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ・給排気ﾌﾞﾛｯｸ組合せﾁｪｯｸ①
                    bolRtn = fncCombinationChk_1(objKtbnStrc, intRightCol, strMsg, strMsgCd)
                    If Not bolRtn Then
                        Exit Try
                    End If

                    '--------------------------------------------------------------
                    '仕切ﾌﾞﾛｯｸ・ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ・給排気ﾌﾞﾛｯｸ組合せﾁｪｯｸ②
                    bolRtn = fncCombinationChk_2(objKtbnStrc, intRightCol, strMsg, strMsgCd)
                    If Not bolRtn Then
                        Exit Try
                    End If
            End Select

            '--------------------------------------------------------------
            '仕切ﾌﾞﾛｯｸ・ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ(3ﾎﾟｰﾄ弁2個内臓形)・給排気ﾌﾞﾛｯｸ組合せﾁｪｯｸ
            If CInt(strUseValues(10)) = 0 And CInt(strUseValues(11)) = 0 Then
                For intR As Integer = 1 To 7
                    strKtbn = strKataValues(intR)
                    intInpUse = CInt(strUseValues(intR))
                    If (strKtbn.IndexOf("66") > -1 _
                    Or strKtbn.IndexOf("67") > -1 _
                    Or strKtbn.IndexOf("76") > -1 _
                    Or strKtbn.IndexOf("77") > -1) _
                    And intInpUse > 0 Then
                        For intR_1 As Integer = 8 To 9
                            strKtbn = strKataValues(intR_1)
                            intInpUse = CInt(strUseValues(intR_1))
                            If intInpUse > 0 _
                            And strKtbn.IndexOf("QK") > -1 Then
                                'message:W2730=３ポート弁２個内臓形には給排気ブロック外部パイロット用（QKまたはQKZ）は選定できません。
                                strMsgCd = "W2730"
                                Exit Try
                            End If
                        Next
                    End If
                Next
            End If

            Return True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncInputChk_2
    '*【処理】
    '*   入力チェック②
    '********************************************************************************************
    Public Shared Function fncInpCheck2(objKtbnStrc As KHKtbnStrc, ByRef dblStdNum As Double, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban
        fncInpCheck2 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            '---------------------------------------------
            '形番選択ﾁｪｯｸ
            For intR As Integer = 0 To 24 - 2
                If intR <> 19 Then
                    If CInt(strUseValues(intR)) > 0 Then
                        If strKataValues(intR).Trim = String.Empty Then
                            If intR <= 13 Then
                                'message:W1400=形番を選択してください。
                                strMsgCd = "W1400"
                                Exit Try
                            Else
                                'message:W1310=添付品の個数を入力したら形番を選択してください。
                                strMsgCd = "W1310"
                                Exit Try
                            End If

                        End If
                    End If
                End If
            Next

            '---------------------------------------------
            '形番要素重複ﾁｪｯｸ
            '- 2～8行目
            For intR As Integer = 1 To 7
                For i As Integer = intR + 1 To 7
                    If strKataValues(intR) = strKataValues(i) _
                    And Not strKataValues(intR) = String.Empty _
                    And Not strKataValues(i) = String.Empty Then
                        'message:W1330=同じ形番が選択されました。
                        strMsgCd = "W1330"
                        Exit Try
                    End If
                Next
            Next

            '- 9～10行目
            If strKataValues(8) = strKataValues(9) _
                And Not strKataValues(8) = String.Empty _
                And Not strKataValues(9) = String.Empty Then
                'message:W1330=同じ形番が選択されました。
                strMsgCd = "W1330"
                Exit Try
            End If

            '- 11～12行目
            If strKataValues(10) = strKataValues(11) _
                And Not strKataValues(10) = String.Empty _
                And Not strKataValues(11) = String.Empty Then
                'message:W1330=同じ形番が選択されました。
                strMsgCd = "W1330"
                Exit Try
            End If

            '- 13～14行目
            If strKataValues(12) = strKataValues(13) _
                And Not strKataValues(12) = String.Empty _
                And Not strKataValues(13) = String.Empty Then
                'message:W1330=同じ形番が選択されました。
                strMsgCd = "W1330"
                Exit Try
            End If

            '---------------------------------------------
            '品名ﾘｽﾄｺﾝﾄﾛｰﾙ 数値ﾃｷｽﾄ入力値ﾁｪｯｸ
            If Not SiyouBLL.fncOtherKataCheck(objKtbnStrc, 15, 23, 20, strMsgCd) Then
                Exit Function
            End If

            '---------------------------------------------
            '設定値判定
            Select Case strSeriesKata
                Case "MN3S0", "MN4S0"
                    If strKataValues(19).ToString.Length <= 0 Then strKataValues(19) = 0
                    If Not SiyouBLL.fncRailchk(strKataValues(20 - 1), strUseValues(20 - 1), dblStdNum, strMsgCd) Then
                        strMsg = "20,0"
                        Exit Function
                    End If
            End Select
            Return True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncCombinationChk_1
    '*【処理】
    '*   '仕切ﾌﾞﾛｯｸ・ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ・給排気ﾌﾞﾛｯｸ組合せﾁｪｯｸ①
    '********************************************************************************************
    Public Shared Function fncCombinationChk_1(objKtbnStrc As KHKtbnStrc, ByVal intRightCol As Integer, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        Dim intR_SA As String = 0
        Dim intParBlk_ALL(0) As Integer  '形番に関わらず選択された列番号を少ない方からｾｯﾄする
        Dim intParBlk_SA(0) As Integer   '形番SAの選択された列番号をｾｯﾄする
        Dim intPos As Integer
        Dim bolValFlg_L As Boolean
        Dim bolValFlg_R As Boolean
        Dim bolExhFlg(0) As Boolean
        Dim bolExhFlg_L As Boolean
        Dim bolExhFlg_R As Boolean
        Dim i As Integer
        Dim r As Integer

        fncCombinationChk_1 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            'SAが選択されている行番号を取得する
            If strUseValues(10) > 0 Then
                If ((Not strKataValues(10).Trim = String.Empty) And strKataValues(10).IndexOf("-SA") > -1) Or _
                   ((Not strKataValues(10).Trim = String.Empty) And strKataValues(10).IndexOf("-SE") > -1) Then
                    intR_SA = 11
                End If
            End If
            If strUseValues(11) > 0 Then
                If ((Not strKataValues(11).Trim = String.Empty) And strKataValues(11).IndexOf("-SA") > -1) Or _
                   ((Not strKataValues(11).Trim = String.Empty) And strKataValues(11).IndexOf("-SE") > -1) Then
                    intR_SA = 12
                End If
            End If

            'SAが選択されていなかった場合ﾁｪｯｸを抜ける
            If intR_SA = 0 Then
                fncCombinationChk_1 = True
                Exit Try
            End If

            '形番に関わらず選択されている仕切ﾌﾞﾛｯｸの列番号を配列にする
            '(0番目:0、最後:最右列番号を設定)
            intParBlk_ALL(0) = 0
            For intC As Integer = 0 To intColCnt - 1
                For intR As Integer = 10 To 11
                    If arySelectInf(intR)(intC) = "1" Then
                        i = i + 1
                        ReDim Preserve intParBlk_ALL(i)
                        intParBlk_ALL(i) = intC
                    End If
                Next
            Next
            ReDim Preserve intParBlk_ALL(i + 1)
            intParBlk_ALL(i + 1) = intRightCol

            'SAのみ選択されている列番号の配列
            '(前後に設定なし)
            r = 0
            For intC As Integer = 0 To intRightCol
                If arySelectInf(intR_SA - 1)(intC) = "1" Then
                    ReDim Preserve intParBlk_SA(r)
                    intParBlk_SA(r) = intC
                    r = r + 1
                End If
            Next
            'SAの行の設置位置が選択されていなかった場合ﾁｪｯｸを抜ける
            If r = 0 Then
                fncCombinationChk_1 = True
                Exit Try
            End If

            '仕切ﾌﾞﾛｯｸが1つ以上選択されていた場合
            If intParBlk_ALL.Length > 2 Then
                '-- SAを1つずつ見ていく
                For intC As Integer = 0 To intParBlk_SA.Length - 1

                    'SAの選択列番号が配列(ALL)の何番目にあるかを取得する
                    For idx_01 As Integer = 0 To intParBlk_ALL.Length - 1
                        If intParBlk_ALL(idx_01) = intParBlk_SA(intC) Then
                            If Not intParBlk_ALL(idx_01) = 0 Then
                                intPos = idx_01
                                Exit For
                            Else
                                intPos = 1  '0列目にﾁｪｯｸされている場合は配列1番目をｾｯﾄ
                                Exit For
                            End If
                        End If
                    Next

                    '==========================================================================
                    '①左右両方向に次の仕切りﾌﾞﾛｯｸがみつかる前にﾊﾞﾙﾌﾞﾌﾞﾛｯｸが存在しなければｴﾗｰ
                    '==========================================================================

                    '配列の前後の列番号との間にﾊﾞﾙﾌﾞﾌﾞﾛｯｸがあるかﾁｪｯｸ
                    '-- 左側
                    bolValFlg_L = False
                    For idx As Integer = intParBlk_ALL(intPos - 1) To intParBlk_ALL(intPos)
                        For intR As Integer = 1 To 7
                            If arySelectInf(intR)(idx) = "1" Then
                                bolValFlg_L = True
                                Exit For
                            End If
                        Next
                        If bolValFlg_L Then
                            Exit For
                        End If
                    Next

                    '-- 右側(左側にあったら)
                    If bolValFlg_L Then
                        bolValFlg_R = False
                        For idx As Integer = intParBlk_ALL(intPos) To intParBlk_ALL(intPos + 1)
                            For intR As Integer = 1 To 7
                                If arySelectInf(intR)(idx) = "1" Then
                                    bolValFlg_R = True
                                    Exit For
                                End If
                            Next
                            If bolValFlg_R Then
                                Exit For
                            End If
                        Next
                    End If

                    '左右両方にﾊﾞﾙﾌﾞﾌﾞﾛｯｸが存在しない場合ｴﾗｰ
                    If Not (bolValFlg_L And bolValFlg_R) Then
                        'message:W2710=仕切ブロックの両側にバルブブロックを指定してください。
                        strMsgCd = "W2710"
                        Exit Try
                    End If

                    '==========================================================================
                    '②左右両方向に次の仕切ﾌﾞﾛｯｸが存在しない場合、給排気ﾌﾞﾛｯｸが存在しないとｴﾗｰ
                    '==========================================================================

                    bolExhFlg_L = False
                    bolExhFlg_R = False
                    '-- 右方向ﾁｪｯｸ
                    If intC < intParBlk_SA.Length - 1 Then
                        For idx As Integer = intParBlk_SA(intC) To intParBlk_SA(intC + 1)
                            For intR As Integer = 8 To 9
                                If strKataValues(intR).IndexOf("-Q-") > -1 _
                                Or strKataValues(intR).IndexOf("-QK-") > -1 _
                                Or strKataValues(intR).IndexOf("-QKZ-") > -1 Then
                                    If arySelectInf(intR)(idx) = "1" Then
                                        bolExhFlg_R = True
                                    End If
                                End If
                            Next
                        Next
                    Else
                        For idx As Integer = intParBlk_SA(intC) To intRightCol
                            For intR As Integer = 8 To 9
                                If strKataValues(intR).IndexOf("-Q-") > -1 _
                                Or strKataValues(intR).IndexOf("-QK-") > -1 _
                                Or strKataValues(intR).IndexOf("-QKZ-") > -1 Then
                                    If arySelectInf(intR)(idx) = "1" Then
                                        bolExhFlg_R = True
                                    End If
                                End If
                            Next
                        Next
                    End If

                    '-- 左方向ﾁｪｯｸ
                    If 0 < intC Then
                        For idx As Integer = intParBlk_SA(intC - 1) To intParBlk_SA(intC)
                            For intR As Integer = 8 To 9
                                If strKataValues(intR).IndexOf("-Q-") > -1 _
                                Or strKataValues(intR).IndexOf("-QK-") > -1 _
                                Or strKataValues(intR).IndexOf("-QKZ-") > -1 Then
                                    If arySelectInf(intR)(idx) = "1" Then
                                        bolExhFlg_L = True
                                    End If
                                End If
                            Next
                        Next
                    Else
                        For idx As Integer = 0 To intParBlk_SA(intC)
                            For intR As Integer = 8 To 9
                                If strKataValues(intR).IndexOf("-Q-") > -1 _
                                Or strKataValues(intR).IndexOf("-QK-") > -1 _
                                Or strKataValues(intR).IndexOf("-QKZ-") > -1 Then
                                    If arySelectInf(intR)(idx) = "1" Then
                                        bolExhFlg_L = True
                                    End If
                                End If
                            Next
                        Next
                    End If

                    If Not (bolExhFlg_L And bolExhFlg_R) Then
                        'message:W2720=仕切ブロックの両側に給排気ブロックを指定してください。
                        strMsgCd = "W2720"
                        Exit Try
                    End If
                Next
            End If
            Return True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncCombinationChk_2
    '*【処理】
    '*   '仕切ﾌﾞﾛｯｸ・ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ・給排気ﾌﾞﾛｯｸ組合せﾁｪｯｸ②
    '********************************************************************************************
    Public Shared Function fncCombinationChk_2(objKtbnStrc As KHKtbnStrc, ByVal intRightCol As Integer, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean

        Dim intParBlk_S(0) As Integer
        Dim intParBlk_SA(0) As Integer
        Dim intR_SA As Integer
        Dim r As Integer
        Dim i As Integer
        Dim bolValFlg_R As Boolean
        Dim bolValFlg_L As Boolean
        Dim intSA_R As Integer
        Dim intSA_L As Integer
        Dim bolExhBlk_R As Boolean
        Dim bolExhBlk_L As Boolean
        Dim bolSA_L As Boolean 'SAが左にあったらTrue
        Dim bolSA_R As Boolean 'SAが右にあったらTrue

        fncCombinationChk_2 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            'S,SP,SEの選択列番号を配列にｾｯﾄ
            '(0番目:0、最後:最右列番号を設定)
            For intC As Integer = 0 To intRightCol
                For intR As Integer = 10 To 11
                    If arySelectInf(intR)(intC) = "1" Then
                        If (Not strKataValues(intR).IndexOf("-SA") > -1) _
                        And (Not strKataValues(intR) = String.Empty) Then
                            'r = r + 1
                            ReDim Preserve intParBlk_S(r)
                            intParBlk_S(r) = intC
                            r = r + 1
                        End If
                    End If
                Next
            Next

            'SAが選択されている行番号を取得する
            If strUseValues(10) > 0 Then
                If (Not strKataValues(10).Trim = String.Empty) _
                And strKataValues(10).IndexOf("-SA") > -1 Then
                    intR_SA = 11
                End If
            End If
            If strUseValues(11) > 0 Then
                If (Not strKataValues(11).Trim = String.Empty) _
                    And strKataValues(11).IndexOf("-SA") > -1 Then
                    intR_SA = 12
                End If
            End If

            'SAのみ選択されている列番号の配列
            '(前後に設定なし)
            If intR_SA <> 0 Then
                i = 0
                For intC As Integer = 0 To intRightCol
                    If arySelectInf(intR_SA - 1)(intC) = "1" Then
                        ReDim Preserve intParBlk_SA(i)
                        intParBlk_SA(i) = intC
                        i = i + 1
                    End If
                Next
            End If

            '===============================================================
            '①左右両方向に次の仕切ﾌﾞﾛｯｸ(SA)または設置位置の端が見つかる前に
            '　ﾊﾞﾙﾌﾞﾌﾞﾛｯｸが存在しないとｴﾗｰ
            '===============================================================

            'S,SP,SEが1つ以上選択されていた場合
            If r > 0 Then
                'If intParBlk_S.Length > 2 Then
                For intC As Integer = 0 To intParBlk_S.Length - 1
                    bolValFlg_R = False
                    bolValFlg_L = False
                    bolSA_R = False
                    bolSA_L = False
                    '-- 右方向ﾁｪｯｸ
                    For idx As Integer = 0 To intParBlk_SA.Length - 1
                        If intParBlk_S(intC) < intParBlk_SA(idx) Then
                            bolSA_R = True
                            intSA_R = idx
                            Exit For
                        End If
                    Next
                    'SAが右にあった場合
                    If bolSA_R Then
                        For idx As Integer = intParBlk_S(intC) To intParBlk_SA(intSA_R)
                            For intR As Integer = 1 To 7
                                If arySelectInf(intR)(idx) = "1" Then
                                    bolValFlg_R = True
                                    Exit For
                                End If
                            Next
                            If bolValFlg_R Then
                                Exit For
                            End If
                        Next
                    Else
                        For idx As Integer = intParBlk_S(intC) To intRightCol
                            For intR As Integer = 1 To 7
                                If arySelectInf(intR)(idx) = "1" Then
                                    bolValFlg_R = True
                                    Exit For
                                End If
                            Next
                            If bolValFlg_R Then
                                Exit For
                            End If
                        Next
                    End If

                    '-- 左方向ﾁｪｯｸ
                    For idx As Integer = intParBlk_SA.Length - 1 To 0 Step -1
                        If intParBlk_SA(idx) <> 0 Then
                            If intParBlk_SA(idx) < intParBlk_S(intC) Then
                                bolSA_L = True
                                intSA_L = idx
                                Exit For
                            End If
                        End If
                    Next
                    'SAが左にあった場合
                    If bolSA_L Then
                        For idx As Integer = intParBlk_S(intC) To intParBlk_SA(intSA_L) Step -1
                            For intR As Integer = 1 To 7
                                If arySelectInf(intR)(idx) = "1" Then
                                    bolValFlg_L = True
                                    Exit For
                                End If
                            Next
                            If bolValFlg_L Then
                                Exit For
                            End If
                        Next
                    Else
                        For idx As Integer = intParBlk_S(intC) To 0 Step -1
                            For intR As Integer = 1 To 7
                                If arySelectInf(intR)(idx) = "1" Then
                                    bolValFlg_L = True
                                    Exit For
                                End If
                            Next
                            If bolValFlg_L Then
                                Exit For
                            End If
                        Next
                    End If

                    If Not (bolValFlg_L And bolValFlg_R) Then
                        'message:W2710=仕切ブロックの両側にバルブブロックを指定してください。
                        strMsgCd = "W2710"
                        Exit Try
                    End If

                    '===============================================================
                    '②左右両方向に次の仕切ﾌﾞﾛｯｸ(SA)または設置位置の端が見つかる前に
                    '　給排気ﾌﾞﾛｯｸが存在しないとｴﾗｰ(ただし片方にあればよい)
                    '===============================================================
                    bolExhBlk_R = False
                    bolExhBlk_L = False
                    bolSA_R = False
                    bolSA_L = False
                    '-- 右方向ﾁｪｯｸ
                    For idx As Integer = 0 To intParBlk_SA.Length - 1
                        If intParBlk_S(intC) < intParBlk_SA(idx) Then
                            bolSA_R = True
                            intSA_R = idx
                            Exit For
                        End If
                    Next
                    'SAが右にあった場合
                    If bolSA_R Then
                        For idx As Integer = intParBlk_S(intC) To intParBlk_SA(intSA_R)
                            For intR As Integer = 8 To 9
                                If strKataValues(intR).IndexOf("-Q-") > -1 _
                                Or strKataValues(intR).IndexOf("-QK-") > -1 _
                                Or strKataValues(intR).IndexOf("-QKZ-") > -1 Then
                                    If arySelectInf(intR)(idx) = "1" Then
                                        bolExhBlk_R = True
                                    End If
                                End If
                            Next
                            If bolExhBlk_R Then
                                Exit For
                            End If
                        Next
                    Else
                        For idx As Integer = intParBlk_S(intC) To intRightCol
                            For intR As Integer = 8 To 9
                                If strKataValues(intR).IndexOf("-Q-") > -1 _
                                Or strKataValues(intR).IndexOf("-QK-") > -1 _
                                Or strKataValues(intR).IndexOf("-QKZ-") > -1 Then
                                    If arySelectInf(intR)(idx) = "1" Then
                                        bolExhBlk_R = True
                                    End If
                                End If
                            Next
                            If bolExhBlk_R Then
                                Exit For
                            End If
                        Next
                    End If

                    '-- 左方向ﾁｪｯｸ
                    For idx As Integer = intParBlk_SA.Length - 1 To 0 Step -1
                        If intParBlk_SA(idx) <> 0 Then
                            If intParBlk_SA(idx) < intParBlk_S(intC) Then
                                bolSA_L = True
                                intSA_L = idx
                                Exit For
                            End If
                        End If
                    Next
                    'SAが左にあった場合
                    If bolSA_L Then
                        For idx As Integer = intParBlk_S(intC) To intParBlk_SA(intSA_L) Step -1
                            For intR As Integer = 8 To 9
                                If strKataValues(intR).IndexOf("-Q-") > -1 _
                                Or strKataValues(intR).IndexOf("-QK-") > -1 _
                                Or strKataValues(intR).IndexOf("-QKZ-") > -1 Then
                                    If arySelectInf(intR)(idx) = "1" Then
                                        bolExhBlk_L = True
                                    End If
                                End If
                            Next
                            If bolExhBlk_L Then
                                Exit For
                            End If
                        Next
                    Else
                        For idx As Integer = intParBlk_S(intC) To 0 Step -1
                            For intR As Integer = 8 To 9
                                If strKataValues(intR).IndexOf("-Q-") > -1 _
                                Or strKataValues(intR).IndexOf("-QK-") > -1 _
                                Or strKataValues(intR).IndexOf("-QKZ-") > -1 Then
                                    If arySelectInf(intR)(idx) = "1" Then
                                        bolExhBlk_L = True
                                    End If
                                End If
                            Next
                            If bolExhBlk_L Then
                                Exit For
                            End If
                        Next
                    End If

                    If (Not bolExhBlk_R) And (Not bolExhBlk_L) Then
                        'message:W2720=仕切ブロックの両側に給･排気ブロック（QまたはQKまたはQKZ）を１ヶ以上を指定してください。
                        strMsgCd = "W2720"
                        Exit Try
                    End If
                Next
            End If
            Return True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try
    End Function
End Class
