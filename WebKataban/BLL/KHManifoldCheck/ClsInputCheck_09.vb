Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_09

    Public Shared intPosRowCnt As Integer = 17
    Public Shared intColCnt As Integer = 25

    Public Shared Function fncInputChk(objKtbnStrc As KHKtbnStrc, ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        fncInputChk = False
        Try
            '入力チェック
            If Not fncInpCheck2(objKtbnStrc, strMsg, strMsgCd) Then
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

    '*   fncInpCheck2
    '*【処理】

    '*   入力チェック2
    '*【引数】

    '*   strKataValues  : 形番の選択値配列          strUseValues    : 使用数の入力値配列
    '*   arySelectInf   : 設置位置の選択値配列      strStdNum       : マニホールド長さ基準
    '********************************************************************************************
    Public Shared Function fncInpCheck2(objKtbnStrc As KHKtbnStrc, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean

        Dim sbCoordinates As New System.Text.StringBuilder
        Dim hshtKataban As New Hashtable
        Dim intChkCnt As Integer

        fncInpCheck2 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            '形番選択チェック
            For intRI As Integer = 0 To intPosRowCnt - 1
                If strKataValues(intRI) = "" And strUseValues(intRI) > 0 Then
                    sbCoordinates.Append(CStr(intRI + 1) & strComma & "0")
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1400"
                    Exit Function
                End If
            Next

            '設置位置コントロール選択チェック
            For intCI As Integer = 0 To intColCnt - 1
                intChkCnt = 0
                For intRI As Integer = Siyou_09.Endb1 - 1 To Siyou_09.MpValve2 - 1
                    '１～１０行目の間で１列２個以上の選択はエラー
                    If arySelectInf(intRI)(intCI) = "1" Then
                        intChkCnt = intChkCnt + 1
                    End If
                Next
                If intChkCnt > 1 Then
                    strMsgCd = "W1390"
                    Exit Function
                End If

                intChkCnt = 0
                For intRI As Integer = Siyou_09.SpReguP - 1 To Siyou_09.ExhaustSp - 1
                    '１１～１５行目の間で１列２個以上の選択はエラー
                    If arySelectInf(intRI)(intCI) = "1" Then
                        intChkCnt = intChkCnt + 1
                    End If
                Next
                If intChkCnt > 1 Then
                    strMsgCd = "W1390"
                    Exit Function
                End If
            Next

            ' 品名リストコントロール 数値テキスト入力値チェック
            If Not SiyouBLL.fncOtherKataCheck(objKtbnStrc, Siyou_09.Silencer1, Siyou_09.Inspect, _
                                     0, strMsgCd) Then
                Exit Function
            End If

            '形番要素重複チェック
            For intRI As Integer = Siyou_09.ElValve1 - 1 To Siyou_09.MpValve2 - 1
                '形番が未選択の場合は省く
                If Len(Trim(strKataValues(intRI))) <> 0 Then
                    If hshtKataban.ContainsKey(strKataValues(intRI)) Then
                        strMsgCd = "W1330"
                        Exit Function
                    End If
                    hshtKataban.Add(strKataValues(intRI), "")
                End If
            Next
            fncInpCheck2 = True
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
    Public Shared Function fncInpCheck1(objKtbnStrc As KHKtbnStrc, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean

        Dim sbCoordinates As New System.Text.StringBuilder
        Dim intChkCnt As Integer = 0
        Dim bolChkFlag As Boolean = False
        Dim intColR As Integer = 0
        Dim intR As Integer = 0
        Dim intMidColR As Integer = 0
        Dim intMidR As Integer = 0
        Dim intMixCon(5) As Integer
        Dim bolMixSwtch(15) As Boolean
        Dim bolMixSwtchMPV As Boolean
        Dim intElectSeq As Integer
        Dim strKataSub As String
        Dim intMixConCnt As Integer = 0
        Dim intMixSwtchCnt As Integer = 0
        Dim intSolenoidCnt As Integer = 0

        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban

        fncInpCheck1 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            '接続位置チェック
            '最右部列取得
            For intCI As Integer = intColCnt To 1 Step -1
                For intRI As Integer = 1 To intPosRowCnt
                    If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                        intR = intRI
                        intColR = intCI
                        Exit For
                    End If
                Next
                If intColR > 0 Then
                    Exit For
                End If
            Next
            '中間列最右部列取得
            For intCI As Integer = intColCnt - 1 To 1 Step -1
                For intRI As Integer = 16 To intPosRowCnt
                    If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                        intMidR = intRI
                        intMidColR = intCI
                        Exit For
                    End If
                Next
            Next

            '列連続チェックエラー
            For intCI As Integer = 1 To intColR
                For intRI As Integer = 1 To Siyou_09.ExhaustSp
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

            'エンドブロック設置位置チェック
            If intR <> Siyou_09.Endb2 Then
                sbCoordinates.Append(CStr(Siyou_09.Endb2) & strComma & CStr(intColR + 1))
                strMsg = sbCoordinates.ToString
                strMsgCd = "W1650"
                Exit Function
            End If
            '中間行が最右に位置する場合もエラー
            If intColR < intMidColR + 1 Then
                sbCoordinates.Append(CStr(intMidR) & strComma & CStr(intMidColR))
                strMsg = sbCoordinates.ToString
                strMsgCd = "W1650"
                Exit Function
            End If


            'エンドブロック複数指定チェック
            If CInt(strUseValues(Siyou_09.Endb1 - 1)) > 1 Then
                sbCoordinates.Append(CStr(Siyou_09.Endb1) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W1100"
                Exit Function
            End If
            If CInt(strUseValues(Siyou_09.Endb2 - 1)) > 1 Then
                sbCoordinates.Append(CStr(Siyou_09.Endb2) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W1100"
                Exit Function
            End If

            '配線ブロック複数指定チェック
            If CInt(strUseValues(Siyou_09.Wiring - 1)) > 1 Then
                sbCoordinates.Append(CStr(Siyou_09.Wiring) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W1530"
                Exit Function
            End If

            '配線ブロック設置位置チェック
            If objKtbnStrc.strcSelection.strOpSymbol(6) = "T10" Then
                Dim strCleaner As String = String.Empty
                Dim strC() As String = objKtbnStrc.strcSelection.strOpSymbol(7).ToString.Split(",")
                For inti As Integer = 0 To strC.Length - 1
                    Select Case strC(inti)
                        Case "CL", "CR"
                            strCleaner = strC(inti).ToString
                    End Select
                Next
                Select Case strCleaner
                    Case "CL"
                        If arySelectInf(Siyou_09.Wiring - 1)(intColR - 2) <> "1" Then
                            strMsgCd = "W2460"
                            Exit Function
                        End If
                    Case "CR"
                        If arySelectInf(Siyou_09.Wiring - 1)(1) <> "1" Then
                            strMsgCd = "W2470"
                            Exit Function
                        End If
                    Case Else
                        If arySelectInf(Siyou_09.Wiring - 1)(intColR - 2) <> "1" Then
                            If arySelectInf(Siyou_09.Wiring - 1)(1) <> "1" Then
                                strMsgCd = "W2480"
                                Exit Function
                            End If
                        End If
                End Select
            Else
                If arySelectInf(Siyou_09.Wiring - 1)(intColR - 2) <> "1" Then
                    If arySelectInf(Siyou_09.Wiring - 1)(1) <> "1" Then
                        strMsgCd = "W2480"
                        Exit Function
                    End If
                End If
            End If

            '電磁弁連数チェック
            For intI As Integer = 0 To UBound(intMixCon)
                intMixCon(intI) = 0
            Next
            For intI As Integer = 0 To UBound(bolMixSwtch)
                bolMixSwtch(intI) = False
            Next
            bolMixSwtchMPV = False
            intElectSeq = 0
            '４～１０行目：バルブブロックチェック
            For intRI As Integer = Siyou_09.ElValve1 - 1 To Siyou_09.MpValve2 - 1

                If CInt(strUseValues(intRI)) > 0 Then
                    If Trim(strKataValues(intRI)).Length > 0 Then
                        If strSeriesKata = "M4TB3" Or strSeriesKata = "M4TB4" Then
                            strKataSub = strKataValues(intRI).Substring(5, 1)
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(1).ToString, 1) = "8" Then
                                Select Case strKataSub
                                    Case "1", "2", "3", "4", "5"
                                        intMixCon(CInt(strKataSub)) = intMixCon(CInt(strKataSub)) + CInt(strUseValues(intRI))
                                End Select

                                If intRI > 7 Then
                                    bolMixSwtchMPV = True
                                End If
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(3).ToString.Substring(1, 1) = "X" Then
                            'ミックスチェック値(切替位置区分)セット
                            strKataSub = strKataValues(intRI).Substring(8, 2)
                            Select Case strKataSub
                                Case "08", "10", "15"
                                    bolMixSwtch(CInt(strKataSub)) = True
                            End Select
                        End If

                        '電磁弁連数値セット
                        intElectSeq = intElectSeq + Int(strUseValues(intRI))
                    End If
                End If
            Next

            If intElectSeq > CInt(objKtbnStrc.strcSelection.strOpSymbol(9).ToString) Then
                strMsgCd = "W1170"
                Exit Function
            End If
            If intElectSeq < CInt(objKtbnStrc.strcSelection.strOpSymbol(9).ToString) Then
                strMsgCd = "W1180"
                Exit Function
            End If

            If strSeriesKata = "M4TB3" Or strSeriesKata = "M4TB4" Then
                If Left(objKtbnStrc.strcSelection.strOpSymbol(1).ToString, 1) = "8" Then
                    For intI As Integer = 0 To UBound(intMixCon)
                        If intMixCon(intI) > 0 Then
                            intMixConCnt = intMixConCnt + 1
                        End If
                    Next
                    If bolMixSwtchMPV = False Then
                        If intMixConCnt < 2 Then
                            strMsgCd = "W1190"
                            Exit Function
                        End If
                    End If
                End If
            End If

            If objKtbnStrc.strcSelection.strOpSymbol(3).ToString.Substring(1, 1) = "X" Then
                For intI As Integer = 0 To UBound(bolMixSwtch)
                    If bolMixSwtch(intI) = True Then
                        intMixSwtchCnt = intMixSwtchCnt + 1
                    End If
                Next
                If intMixSwtchCnt < 2 Then
                    strMsgCd = "W1220"
                    Exit Function
                End If
            End If

            'ソレノイドカウントチェック
            If strSeriesKata = "M4TB3" Or strSeriesKata = "M4TB4" Then
                For intRI As Integer = Siyou_09.ElValve1 - 1 To Siyou_09.MpValve2 - 1
                    If CInt(strUseValues(intRI)) > 0 Then
                        If strKataValues(intRI).Substring(5, 1) = "1" Then
                            intSolenoidCnt = intSolenoidCnt + 1 * CInt(strUseValues(intRI))
                        ElseIf strKataValues(intRI).Substring(5, 1) = "-" Then
                        Else
                            intSolenoidCnt = intSolenoidCnt + 2 * CInt(strUseValues(intRI))
                        End If
                    End If
                Next
                If intSolenoidCnt > KHKataban.fncGetMaxSol(objKtbnStrc.strcSelection.strOpSymbol, 9) Then
                    strMsgCd = "W1150"
                    Exit Function
                End If
            End If

            strMsg = String.Empty
            'バルブブロック・レギュレーター・スペーサ組み合わせチェック
            For intCI As Integer = 0 To intColCnt - 1
                For intRI As Integer = Siyou_09.SpReguP - 1 To Siyou_09.ExhaustSp - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        intChkCnt = 0
                        For intRI2 As Integer = Siyou_09.ElValve1 - 1 To Siyou_09.MpValve2 - 1
                            If arySelectInf(intRI2)(intCI) = "1" Then
                                intChkCnt = intChkCnt + 1
                            End If
                        Next
                        If intChkCnt = 0 Then
                            strMsg &= intRI + 1 & "," & intCI + 1
                            strMsgCd = "W4090"
                            Exit Function
                        End If
                    End If
                Next
            Next

            '仕切プラグ(給気用)・単独給気スペーサ組み合わせチェック
            For intCI As Integer = 0 To intColCnt - 2
                If arySelectInf(Siyou_09.PartitionS - 1)(intCI) = "1" Then
                    '仕切プラグ(給気用)が選択されている場合、右の列について確認
                    intChkCnt = 0
                    bolChkFlag = False
                    For intCI2 As Integer = intCI + 1 To intColCnt - 2
                        '単独給気スペーサが見つかる前に次の仕切プラグ(給気用)が見つかったらエラー
                        If arySelectInf(Siyou_09.SupplySp - 1)(intCI2) = "1" Then
                            intChkCnt = intChkCnt + 1
                        End If
                        If arySelectInf(Siyou_09.PartitionS - 1)(intCI2) = "1" Then
                            bolChkFlag = True
                            Exit For
                        End If
                    Next
                    If bolChkFlag = True And intChkCnt = 0 Then
                        strMsgCd = "W4120"
                        Exit Function
                    End If
                End If
            Next

            '仕切プラグ(排気用)・単独排気スペーサ組み合わせチェック
            For intCI As Integer = 0 To intColCnt - 2
                If arySelectInf(Siyou_09.PartitionE - 1)(intCI) = "1" Then
                    '仕切プラグ(排気用)が選択されている場合、右の列について確認
                    intChkCnt = 0
                    bolChkFlag = False
                    For intCI2 As Integer = intCI + 1 To intColCnt - 2
                        '単独排気スペーサが見つかる前に次の仕切プラグ(排気用)が見つかったらエラー
                        If arySelectInf(Siyou_09.ExhaustSp - 1)(intCI2) = "1" Then
                            intChkCnt = intChkCnt + 1
                        End If
                        If arySelectInf(Siyou_09.PartitionE - 1)(intCI2) = "1" Then
                            bolChkFlag = True
                            Exit For
                        End If
                    Next
                    If bolChkFlag = True And intChkCnt = 0 Then
                        strMsgCd = "W2490"
                        Exit Function
                    End If
                End If
            Next
            fncInpCheck1 = True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try

    End Function
End Class
