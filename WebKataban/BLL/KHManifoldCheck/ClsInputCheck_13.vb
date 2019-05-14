Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_13

    Public Shared intPosRowCnt As Integer = 17
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
            If Not fncInpCheck1(objKtbnStrc, strMsg, strMsgCd) Then
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

    '*   fncInpCheck
    '*【処理】

    '*   入力チェック
    '*【引数】

    '*   strKataValues  : 形番の選択値配列          strUseValues    : 使用数の入力値配列
    '*   arySelectInf   : 設置位置の選択値配列      arySelectInf    ：設置位置の選択値配列(中間行)
    '********************************************************************************************
    Public Shared Function fncInpCheck1(objKtbnStrc As KHKtbnStrc, ByRef strMsg As String, ByRef strMsgCd As String) As Boolean

        Dim sbCoordinates As New System.Text.StringBuilder
        Dim bolPosChkFlg As Boolean = False
        Dim intColR As Integer = 0
        Dim intColRPlug As Integer = 0
        Dim intSolenoidCnt As Integer = 0
        Dim bolExhaustFlg1 As Boolean = False
        Dim bolExhaustFlg2 As Boolean = False
        Dim intMixCon(5) As Integer
        Dim bolMixConMpv As Boolean
        Dim bolMixSwtch(10) As Boolean
        Dim intElectSeq As Integer
        Dim intMixConCnt As Integer = 0
        Dim intMixSwtchCnt As Integer = 0
        Dim strKataSub As String
        Dim strkatasub2 As String

        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban

        fncInpCheck1 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            Dim strConCaliber As String = String.Empty
            Dim intElectMax As Integer = 0
            If strSeriesKata = "MN4TB1" Or strSeriesKata = "MN4TB2" Then
                strConCaliber = objKtbnStrc.strcSelection.strOpSymbol(3).ToString
                intElectMax = CLng(objKtbnStrc.strcSelection.strOpSymbol(9).ToString)
            Else
                strConCaliber = objKtbnStrc.strcSelection.strOpSymbol(1).ToString
                intElectMax = CLng(objKtbnStrc.strcSelection.strOpSymbol(7).ToString)
            End If

            '接続位置チェック
            For intCI As Integer = intColCnt To 1 Step -1
                bolPosChkFlg = False
                For intRI As Integer = 1 To intPosRowCnt
                    '仕切プラグを含まない最右列取得
                    If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                        If intColR = 0 Then
                            intColR = intCI
                        End If
                        bolPosChkFlg = True
                    End If
                    '仕切プラグの最右列取得
                    If intCI < intColCnt Then
                        If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                            If intColRPlug = 0 Then
                                intColRPlug = intCI
                            End If
                        End If
                    End If
                Next

                '連続未チェックエラー
                If intColR > 0 And bolPosChkFlg = False Then
                    sbCoordinates.Append("0" & strComma & CStr(intCI))
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1020"
                    Exit Function
                End If
            Next

            '最右列再セット
            If intColR < intColRPlug Then
                intColR = intColRPlug
            End If

            '未チェックエラー
            If intColR = 0 And intColRPlug = 0 Then
                strMsgCd = "W1030"
                Exit Function
            End If

            '左側エンドプレートチェック
            If Trim(strKataValues(Siyou_13.Wiring - 1)).Length = 0 Then
                strMsgCd = "W1400"
                Exit Function
            End If

            'エンドプレートチェック
            If arySelectInf(0)(0) = "1" Then        '配線ブロック左側仕様
                'エンドブロック(３行目)が最右
                If arySelectInf(Siyou_13.End2 - 1)(intColR - 1) = "1" Then
                Else
                    sbCoordinates.Append(CStr(Siyou_13.End2) & strComma & CStr(intColR + 1))
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1410"
                    Exit Function
                End If
            Else                                    '配線ブロック右側仕様
                '配線ブロックが最右
                If arySelectInf(Siyou_13.Wiring - 1)(intColR - 1) = "1" Then
                    '配線ブロック左隣がエンドブロック(３行目)
                    If arySelectInf(Siyou_13.End2 - 1)(intColR - 2) = "1" Then
                    Else
                        sbCoordinates.Append("0" & strComma & CStr(intColR - 1))
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W1520"
                        Exit Function
                    End If
                Else
                    '配線ブロック最右以外
                    strMsgCd = "W1510"
                    Exit Function
                End If
            End If


            '配線ブロック複数入力チェック
            If CInt(strUseValues(Siyou_13.Wiring - 1)) > 1 Then
                sbCoordinates.Append(CStr(Siyou_13.Wiring) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W1530"
                Exit Function
            End If

            'エンドブロックブロック複数入力チェック
            If CInt(strUseValues(Siyou_13.End1 - 1)) > 1 Then
                sbCoordinates.Append(CStr(Siyou_13.End1) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W2310"
                Exit Function
            End If
            If CInt(strUseValues(Siyou_13.End2 - 1)) > 1 Then
                sbCoordinates.Append(CStr(Siyou_13.End2) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W2320"
                Exit Function
            End If

            'エンドプレート形番チェック
            If arySelectInf(0)(0) = "1" Then        '配線ブロック左側仕様
                'エンドブロック２行目２列目未チェック
                If arySelectInf(1)(1) = "1" Then
                Else
                    sbCoordinates.Append(CStr(Siyou_13.End1) & strComma & CStr(Siyou_13.End1))
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1540"
                    Exit Function
                End If

                'エンドブロック形番チェック
                If strSeriesKata = "MN4TB2" Then
                    If Left(strKataValues(Siyou_13.End1 - 1), 5) = "2NEA1" Then
                    Else
                        strMsgCd = "W1550"
                        Exit Function
                    End If
                    If Left(strKataValues(Siyou_13.End2 - 1), 5) = "2NEB2" Then
                    Else
                        strMsgCd = "W1560"
                        Exit Function
                    End If
                Else
                    If Left(strKataValues(Siyou_13.End1 - 1), 5) = "1NEA1" Then
                    Else
                        strMsgCd = "W1550"
                        Exit Function
                    End If
                    If Left(strKataValues(Siyou_13.End2 - 1), 5) = "1NEB2" Then
                    Else
                        strMsgCd = "W1560"
                        Exit Function
                    End If
                End If
            Else                                    '配線ブロック右側仕様
                If strSeriesKata = "MN4TB2" Then
                    If Left(strKataValues(Siyou_13.End1 - 1), 5) = "2NEB1" Then
                    Else
                        strMsgCd = "W1550"
                        Exit Function
                    End If
                    If Left(strKataValues(Siyou_13.End2 - 1), 5) = "2NEA2" Then
                    Else
                        strMsgCd = "W1560"
                        Exit Function
                    End If
                Else
                    If Left(strKataValues(Siyou_13.End1 - 1), 5) = "1NEB1" Then
                    Else
                        strMsgCd = "W1550"
                        Exit Function
                    End If
                    If Left(strKataValues(Siyou_13.End2 - 1), 5) = "1NEA2" Then
                    Else
                        strMsgCd = "W1560"
                        Exit Function
                    End If
                End If
            End If

            'ソレノイドカウントチェック
            For intRI As Integer = Siyou_13.Valve1 - 1 To Siyou_13.Valve6 - 1
                If CInt(strUseValues(intRI)) > 0 Then
                    If Trim(strKataValues(intRI)).Length > 0 Then
                        If strKataValues(intRI).Substring(5, 1) = "1" Then
                            intSolenoidCnt = intSolenoidCnt + 1 * CInt(strUseValues(intRI))
                        ElseIf strKataValues(intRI).Substring(5, 1) = "-" Then
                        Else
                            intSolenoidCnt = intSolenoidCnt + 2 * CInt(strUseValues(intRI))
                        End If
                    End If
                    If intSolenoidCnt > KHKataban.fncGetMaxSol(objKtbnStrc.strcSelection.strOpSymbol, 13) Then
                        strMsgCd = "W1150"
                        Exit Function
                    End If
                End If
            Next

            '給排気チェック
            For intCI As Integer = 0 To intColR - 1
                For intRI As Integer = Siyou_13.Exhaust1 - 1 To Siyou_13.Exhaust2 - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        bolExhaustFlg1 = True
                        bolExhaustFlg2 = True
                    End If
                Next
                For intRI As Integer = Siyou_13.Exhaust3 - 1 To Siyou_13.Exhaust4 - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        bolExhaustFlg1 = True
                    End If
                Next
                For intRI As Integer = Siyou_13.Exhaust5 - 1 To Siyou_13.Exhaust6 - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        bolExhaustFlg2 = True
                    End If
                Next

                'エンドブロックチェック
                If arySelectInf(Siyou_13.End2 - 1)(intCI) = "1" Then
                    If bolExhaustFlg1 = True And bolExhaustFlg2 = True Then
                        bolExhaustFlg1 = False
                        bolExhaustFlg2 = False
                    Else
                        sbCoordinates.Append(Siyou_13.Exhaust1 & strComma & "0|" & Siyou_13.Exhaust2 & strComma & "0|" & _
                                             Siyou_13.Exhaust3 & strComma & "0|" & Siyou_13.Exhaust4 & strComma & "0|" & _
                                             Siyou_13.Exhaust5 & strComma & "0|" & Siyou_13.Exhaust6 & strComma & "0|")
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W1420"
                        Exit Function
                    End If
                End If

                '仕切プラグチェック
                If intCI < intColR - 2 Then
                    If arySelectInf(Siyou_13.Partition1 - 1)(intCI) = "1" And _
                       arySelectInf(Siyou_13.Partition2 - 1)(intCI) = "1" Then
                        If bolExhaustFlg1 = True And bolExhaustFlg2 = True Then
                            bolExhaustFlg1 = False
                            bolExhaustFlg2 = False
                        Else
                            sbCoordinates.Append("0" & strComma & CStr(intCI + 1))
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W1430"
                            Exit Function
                        End If
                    ElseIf arySelectInf(Siyou_13.Partition1 - 1)(intCI) = "1" Then
                        If bolExhaustFlg1 = True Then
                            bolExhaustFlg1 = False
                        Else
                            sbCoordinates.Append("0" & strComma & CStr(intCI + 1))
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W1430"
                            Exit Function
                        End If
                    ElseIf arySelectInf(Siyou_13.Partition2 - 1)(intCI) = "1" Then
                        If bolExhaustFlg2 = True Then
                            bolExhaustFlg2 = False
                        Else
                            sbCoordinates.Append("0" & strComma & CStr(intCI + 1))
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W1430"
                            Exit Function
                        End If
                    End If
                End If
            Next

            ''********** 電磁弁連数チェック ******************************
            For intI As Integer = 0 To UBound(intMixCon)
                intMixCon(intI) = 0
            Next
            For intI As Integer = 0 To UBound(bolMixSwtch)
                bolMixSwtch(intI) = False
            Next
            bolMixConMpv = False
            intElectSeq = 0
            '４～９行目：バルブブロックチェック
            For intRI As Integer = Siyou_13.Valve1 - 1 To Siyou_13.Valve6 - 1

                If CInt(strUseValues(intRI)) > 0 Then
                    If Trim(strKataValues(intRI)).Length > 0 Then
                        If strSeriesKata = "MN4TB1" Or strSeriesKata = "MN4TB2" Then
                            strKataSub = strKataValues(intRI).Substring(5, 1)
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(1).ToString, 1) = "8" Then
                                If strKataSub >= "1" And strKataSub <= "5" Then
                                    intMixCon(CInt(strKataSub)) = intMixCon(CInt(strKataSub)) + CInt(strUseValues(intRI))
                                End If
                                'ミックスチェック値(接続口径)セット
                                If strKataValues(intRI).Substring(7, 3) = "MPV" Then
                                    bolMixConMpv = True
                                End If
                            End If
                        Else
                            strKataSub = strKataValues(intRI).Substring(4, 1)
                            'ミックスチェック値(接続口径)セット
                            If CInt(strKataSub) >= 1 And CInt(strKataSub) <= 2 Then
                                intMixCon(CInt(strKataSub)) = CInt(strUseValues(intRI))
                            End If
                        End If

                        If strConCaliber.ToString.Substring(1, 1) = "X" Then
                            'ミックスチェック値(切替位置区分)セット
                            strKataSub = strKataValues(intRI).Substring(9, 1)
                            strkatasub2 = Left(strKataValues(intRI) & Space(12), 12).Substring(11, 1)
                            If strKataSub = "4" Or strKataSub = "6" Or strKataSub = "8" Then
                                bolMixSwtch(CInt(strKataSub)) = True
                            ElseIf strkatasub2 = "4" Or strkatasub2 = "6" Or strkatasub2 = "8" Then
                                bolMixSwtch(CInt(strkatasub2)) = True
                            End If
                            If strKataSub = "1" Or strkatasub2 = "1" Then
                                bolMixSwtch(10) = True
                            End If
                        End If
                    End If

                    '電磁弁連数値セット
                    intElectSeq = intElectSeq + Int(strUseValues(intRI))
                End If
            Next

            '電磁弁連数エラーチェック
            If intElectSeq > intElectMax Then
                strMsgCd = "W1170"
                Exit Function
            End If
            If intElectSeq < intElectMax Then
                strMsgCd = "W1180"
                Exit Function
            End If

            'ミックスチェック(接続口径)チェック
            If strSeriesKata = "MN4TB1" Or strSeriesKata = "MN4TB2" Then
                If Left(objKtbnStrc.strcSelection.strOpSymbol(1).ToString, 1) = "8" Then
                    For intI As Integer = 0 To UBound(intMixCon)
                        If intMixCon(intI) > 0 Then
                            intMixConCnt = intMixConCnt + 1
                        End If
                    Next
                    If bolMixConMpv = True Then
                        If intMixConCnt < 1 Then
                            strMsgCd = "W1190"
                            Exit Function
                        End If
                    Else
                        If intMixConCnt <= 1 Then
                            strMsgCd = "W1190"
                            Exit Function
                        End If
                    End If
                End If
            Else
                If intMixCon(1) = 0 Or intMixCon(2) = 0 Then
                    strMsgCd = "W1570"
                    Exit Function
                End If
            End If

            'ミックスチェック(切替位置区分)チェック
            If strConCaliber.Substring(1, 1) = "X" Then
                For intI As Integer = 0 To UBound(bolMixSwtch)
                    If bolMixSwtch(intI) = True Then
                        intMixSwtchCnt = intMixSwtchCnt + 1
                    End If
                Next
                If intMixSwtchCnt <= 1 Then
                    strMsgCd = "W1580"
                    Exit Function
                End If
            End If

            fncInpCheck1 = True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try

    End Function

    '********************************************************************************************
    '*【関数名】

    '*   fncInpCheck2
    '*【処理】

    '*   入力チェック2
    '*【引数】

    '*   strKataValues  : 形番の選択値配列          strUseValues    : 使用数の入力値配列
    '*   arySelectInf   : 設置位置の選択値配列      strStdNum       : マニホールド長さ基準
    '********************************************************************************************
    Public Shared Function fncInpCheck2(objKtbnStrc As KHKtbnStrc, ByRef dblStdNum As Double, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        Dim sbCoordinates As New System.Text.StringBuilder
        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban

        fncInpCheck2 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            ''形番未選択チェック
            For intRI As Integer = 0 To intPosRowCnt - 1
                If strUseValues(intRI) > 0 And strKataValues(intRI) = "" Then
                    sbCoordinates.Append(CStr(intRI + 1) & strComma & "0")
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1400"
                    Exit Function
                End If
            Next

            ''形番リスト重複チェック
            If Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_13.Valve1, Siyou_13.Valve6) Or _
               Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_13.Exhaust1, Siyou_13.Exhaust2) Or _
               Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_13.Exhaust3, Siyou_13.Exhaust4) Or _
               Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_13.Exhaust5, Siyou_13.Exhaust6) Or _
               Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_13.Silencer1, Siyou_13.Silencer4) Then
                strMsgCd = "W1330"
                Exit Function
            End If


            ''取付レール長さ設定値チェック
            If strKataValues(Siyou_13.Rail - 1).ToString.Length <= 0 Then strKataValues(Siyou_13.Rail - 1) = 0
            If Not SiyouBLL.fncRailchk(strKataValues(Siyou_13.Rail - 1), strUseValues(Siyou_13.Rail - 1), dblStdNum, strMsgCd) Then
                strMsg = Siyou_13.Rail & ",0"
                Exit Function
            End If

            'ブランクプラグ個数チェック
            If Not SiyouBLL.fncOtherKataCheck(objKtbnStrc, Siyou_13.Silencer1, Siyou_13.Cable2, Siyou_13.Rail, strMsgCd) Then
                Exit Function
            End If

            fncInpCheck2 = True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try

    End Function
End Class
