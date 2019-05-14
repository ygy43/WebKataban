Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_02

    Public Shared intColCnt As Integer = 40         'RM1803032_マニホールド連数拡張
    Public Shared intPosRowCnt As Integer = 16

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
    '*   arySelectInf   : 設置位置の選択値配列
    '********************************************************************************************
    Public Shared Function fncInpCheck1(objKtbnStrc As KHKtbnStrc, ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        Dim sbCoordinates As New System.Text.StringBuilder
        Dim bolRightChkFlg As Boolean = False
        Dim bolPosChkFlg As Boolean = False
        Dim bolExhaustFlg1 As Boolean = False
        Dim bolExhaustFlg2 As Boolean = False
        Dim intMixCon(5) As Integer
        Dim bolMixSwtch(10) As Boolean
        Dim intMixConCnt As Integer = 0
        Dim intMixSwtchCnt As Integer = 0
        Dim intElectSeq As Integer
        Dim strKataban As String
        Dim strKataSub As String

        Dim strSeriesKata As String = String.Empty      '機種
        strSeriesKata = objKtbnStrc.strcSelection.strSeriesKataban

        fncInpCheck1 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            ''左側エンドプレートチェック
            If strKataValues(0) = "" Then
                strMsgCd = "W1400"
                Exit Function
            End If

            ''接続位置チェック
            For intCI As Integer = intColCnt To 1 Step -1
                bolPosChkFlg = False

                For intRI As Integer = 1 To intPosRowCnt
                    If bolRightChkFlg = False Then
                        If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                            If intRI = Siyou_02.End2 Then
                                bolRightChkFlg = True
                                bolPosChkFlg = True
                            Else
                                sbCoordinates.Append("0" & strComma & CStr(intCI))
                                strMsg = sbCoordinates.ToString
                                strMsgCd = "W1410"
                                Exit Function
                            End If
                        End If
                    Else
                        If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                            bolPosChkFlg = True
                        End If
                    End If
                Next

                '未チェックエラー
                If bolRightChkFlg = True And bolPosChkFlg = False Then
                    sbCoordinates.Append("0" & strComma & CStr(intCI))
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1020"
                    Exit Function
                End If
            Next

            ''********** 給排気チェック ******************************
            For intCI As Integer = 0 To intColCnt - 1
                ''給排気フラグセット
                If arySelectInf(Siyou_02.Exhaust1 - 1)(intCI) = "1" Or _
                   arySelectInf(Siyou_02.Exhaust2 - 1)(intCI) = "1" Then
                    bolExhaustFlg1 = True
                    bolExhaustFlg2 = True
                End If
                If arySelectInf(Siyou_02.Exhaust3 - 1)(intCI) = "1" Or _
                   arySelectInf(Siyou_02.Exhaust4 - 1)(intCI) = "1" Then
                    bolExhaustFlg1 = True
                End If
                If arySelectInf(Siyou_02.Exhaust5 - 1)(intCI) = "1" Or _
                   arySelectInf(Siyou_02.Exhaust6 - 1)(intCI) = "1" Then
                    bolExhaustFlg2 = True
                End If

                ''２行目：エンドブロックチェック
                If arySelectInf(Siyou_02.End2 - 1)(intCI) = "1" Then
                    If bolExhaustFlg1 = True And bolExhaustFlg2 = True Then
                        bolExhaustFlg1 = False
                        bolExhaustFlg2 = False
                    Else
                        sbCoordinates.Append("0" & strComma & CStr(intCI + 1))
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W1420"
                        Exit Function
                    End If
                End If

                ''１５行目：仕切りブロックチェック
                If arySelectInf(Siyou_02.Partition1 - 1)(intCI) = "1" Then
                    If bolExhaustFlg1 = True And bolExhaustFlg2 = True Then
                        bolExhaustFlg1 = False
                        bolExhaustFlg2 = False
                    Else
                        sbCoordinates.Append("0" & strComma & CStr(intCI + 1))
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W1430"
                        Exit Function
                    End If
                End If

                ''１６行目：仕切りブロックチェック
                If arySelectInf(Siyou_02.Partition2 - 1)(intCI) = "1" Then
                    If bolExhaustFlg2 = True Then
                        bolExhaustFlg2 = False
                    Else
                        sbCoordinates.Append("0" & strComma & CStr(intCI + 1))
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W1430"
                        Exit Function
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
            intElectSeq = 0
            '９～１４行目：電磁弁付バルブブロックチェック
            For intRI As Integer = Siyou_02.Valve1 - 1 To Siyou_02.Valve6 - 1
                strKataban = Trim(strKataValues(intRI))
                '形番要素が選択されている場合
                If Len(strKataban) > 0 Then
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(1).ToString, 1) = "8" Then
                        'ミックスチェック値(接続口径)セット
                        strKataSub = strKataban.Substring(5, 1)
                        If CInt(strKataSub) >= 1 And CInt(strKataSub) <= 5 Then
                            intMixCon(CInt(strKataSub)) = intMixCon(CInt(strKataSub)) + CInt(strUseValues(intRI))
                        End If
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(2).ToString.Substring(1, 1) = "X" And _
                        CInt(strUseValues(intRI)) > 0 Then
                        'ミックスチェック値(切替位置区分)セット
                        If strSeriesKata = "MN4KB1" Then
                            strKataSub = strKataban.Substring(10, 1)
                            If strKataSub = "4" Or strKataSub = "6" Or strKataSub = "8" Then
                                bolMixSwtch(CInt(strKataSub)) = True
                            End If
                        Else
                            strKataSub = strKataban.Substring(9, 1)
                            If strKataSub = "6" Or strKataSub = "8" Then
                                bolMixSwtch(CInt(strKataSub)) = True
                            End If
                            If strKataSub = "1" Then
                                bolMixSwtch(10) = True
                            End If
                        End If
                    End If

                    '電磁弁連数値セット
                    If CInt(strUseValues(intRI)) > 0 Then
                        intElectSeq = intElectSeq + Int(strUseValues(intRI))
                    End If
                End If
            Next

            'エラーチェック
            If intElectSeq > objKtbnStrc.strcSelection.strOpSymbol(7) Then
                strMsgCd = "W1170"
                Exit Function
            End If

            If intElectSeq < objKtbnStrc.strcSelection.strOpSymbol(7) Then
                strMsgCd = "W1180"
                Exit Function
            End If

            If Left(objKtbnStrc.strcSelection.strOpSymbol(1).ToString, 1) = "8" Then
                For intI As Integer = 0 To UBound(intMixCon)
                    If intMixCon(intI) > 0 Then
                        intMixConCnt = intMixConCnt + 1
                    End If
                Next
                If intMixConCnt <= 1 Then
                    strMsgCd = "W1190"
                    Exit Function
                End If
            End If

            If objKtbnStrc.strcSelection.strOpSymbol(2).ToString.Substring(1, 1) = "X" Then
                For intI As Integer = 0 To UBound(bolMixSwtch)
                    If bolMixSwtch(intI) = True Then
                        intMixSwtchCnt = intMixSwtchCnt + 1
                    End If
                Next
                If intMixSwtchCnt <= 1 Then
                    strMsgCd = "W1440"
                    Exit Function
                End If
            End If

            ''*********** サイレンサ・ブランクプラグ・検査成績書チェック *********************************
            If Not SiyouBLL.fncOtherKataCheck(objKtbnStrc, Siyou_02.Silencer1, Siyou_02.Inspect, Siyou_02.Rail, strMsgCd) Then
                Exit Function
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
    Public Shared Function fncInpCheck2(objKtbnStrc As KHKtbnStrc, dblStdNum As Double, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        Dim sbCoordinates As New System.Text.StringBuilder

        fncInpCheck2 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            ''形番未選択チェック
            For intRI As Integer = 0 To intPosRowCnt - 1
                If strKataValues(intRI) = "" Then
                    For intCI As Integer = 0 To intColCnt - 1
                        If arySelectInf(intRI)(intCI) = "1" Then
                            sbCoordinates.Append(CStr(intRI + 1) & strComma & "0")
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W1400"
                            Exit Function
                        End If
                    Next
                End If
            Next

            ''エンドブロック複数不可チェック
            If CInt(strUseValues(Siyou_02.End1 - 1)) > 1 Then
                sbCoordinates.Append(CStr(Siyou_02.End1) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W2320"
                Exit Function
            End If
            If CInt(strUseValues(Siyou_02.End2 - 1)) > 1 Then
                sbCoordinates.Append(CStr(Siyou_02.End2) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W2320"
                Exit Function
            End If

            ''品名リストコントロール　形番リスト重複チェック
            If Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_02.Exhaust1, Siyou_02.Exhaust2) Or _
               Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_02.Exhaust3, Siyou_02.Exhaust4) Or _
               Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_02.Exhaust5, Siyou_02.Exhaust6) Or _
               Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_02.Valve1, Siyou_02.Valve6) Or _
               Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_02.Silencer1, Siyou_02.Silencer2) Or _
               Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_02.Plug1, Siyou_02.Plug2) Then
                strMsgCd = "W1330"
                Exit Function
            End If

            ''取付レール長さ設定値チェック
            If strKataValues(Siyou_02.Rail - 1).ToString.Length <= 0 Then strKataValues(Siyou_02.Rail - 1) = 0
            If Not SiyouBLL.fncRailchk(strKataValues(Siyou_02.Rail - 1), CDbl(strUseValues(Siyou_02.Rail - 1)), dblStdNum, strMsgCd) Then
                strMsg = Siyou_02.Rail & ",0"
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
