Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_16

    Private Const CST_BLANK As String = ""
    Private Const CST_ZERO As String = "0"

    Public Shared intPosRowCnt As Integer = 23
    Private Const CST_ROW As Integer = 18
    Public Shared intColCnt As Integer = 40          'RM1803032_マニホールド連数拡張

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
    '*   fncInputChk①
    '*【処理】
    '*   入力内容をチェックする
    '*【更新】
    '*
    '********************************************************************************************
    Public Shared Function fncInpCheck1(objKtbnStrc As KHKtbnStrc, ByRef strMsg As String, ByRef strMsgCd As String) As Boolean

        Dim intMixChk(8) As Integer      '0番目""

        Dim intRightCol As Integer = -1
        Dim strXY As String = ""
        Dim strXY_R As String = ""
        Dim strXY_L As String = ""
        Dim bolFlg1 As Boolean
        Dim intFlg2_R As Integer
        Dim intFlg2_L As Integer
        Dim intSoreCnt As Integer = 0
        Dim intCnt4 As Integer = 0
        Dim intCnt5 As Integer = 0
        Dim intTtlCnt5 As Integer = 0
        Dim intFlg5 As Integer = 0
        Dim intCnt6_1 As Integer = 0
        Dim intCnt6_2 As Integer = 0
        Dim intCnt6_3 As Integer = 0
        Dim intCnt7 As Integer = 0
        Dim strFlg8_1(intColCnt - 1) As String
        Dim strFlg8_2(intColCnt - 1) As String
        Dim intCnt8_1 As Integer = 0
        Dim intCnt8_2 As Integer = 0
        Dim strFlg9_1(intColCnt - 1) As String
        Dim strFlg9_2(intColCnt - 1) As String
        Dim intCnt9_1 As Integer = 0
        Dim intCnt9_2 As Integer = 0

        Dim intPosCnt As Integer
        Dim intEndPosCnt As Integer
        Dim intWirPosCnt As Integer
        Dim intValPosCnt As Integer
        Dim intPrtPosCnt As Integer
        Dim intSpcPosCnt As Integer

        Dim strCoordinates As String

        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban
        Dim strKeyKata As String = objKtbnStrc.strcSelection.strKeyKataban

        fncInpCheck1 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            '設置位置選択情報取得

            '-- 選択されているｾﾙのうち最右列を取得する
            For intC As Integer = intColCnt - 1 To 0 Step -1
                bolFlg1 = False
                '-中間行
                If Not intC = intColCnt - 1 Then
                    For intR As Integer = Siyou_16.Partition1 - 1 To Siyou_16.Partition2 - 1
                        If arySelectInf(intR)(intC) = "1" Then
                            bolFlg1 = True
                            Exit For
                        End If
                    Next
                    If bolFlg1 Then
                        intRightCol = intC
                        Exit For
                    End If
                End If
                '-通常行
                For intR As Integer = 0 To Siyou_16.RegulatorB - 1
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

            '---------------------------------------------
            '①未接続位置ﾁｪｯｸ

            '-- 選択されているｾﾙが全くない場合
            If intRightCol = 0 Then
                'message:W1030=設置位置が未入力です。選択してください。
                strMsgCd = "W1030"
                Exit Try
            End If

            '-- 選択されているｾﾙがない列がある場合
            strXY = CST_BLANK
            For intC As Integer = 0 To intRightCol
                bolFlg1 = False
                For intR As Integer = 0 To Siyou_16.RegulatorB - 1
                    If arySelectInf(intR)(intC) = "1" Then
                        bolFlg1 = True
                        Exit For
                    End If
                Next
                If Not bolFlg1 Then
                    'message:W1020=選択されない接続位置があります。接続位置=[1]
                    For idx As Integer = 1 To Siyou_16.MPValve2
                        strXY = strXY & CStr(idx) & strComma & CStr(intC + 1) & strPipe
                    Next
                    strMsg = strXY
                    strMsgCd = "W1020"
                    Exit Try
                End If
            Next

            '******* 入出力ブロックチェック(7.2 / 7.3) ******************************
            If Int(strUseValues(Siyou_16.InOut1 - 1)) > 0 And _
               Len(Trim(strKataValues(Siyou_16.InOut1 - 1))) = 0 Then
                strCoordinates = CStr(Siyou_16.InOut1) & strComma & "0"
                strMsg = strCoordinates
                strMsgCd = "W1600"
                Exit Try
            End If
            If Int(strUseValues(Siyou_16.InOut2 - 1)) > 0 And _
               Len(Trim(strKataValues(Siyou_16.InOut2 - 1))) = 0 Then
                strCoordinates = CStr(Siyou_16.InOut2) & strComma & "0"
                strMsg = strCoordinates
                strMsgCd = "W1610"
                Exit Try
            End If

            '-- 列ごとのチェック(1列のなかでエンドブロック/電磁弁＆マスキングプレート/スペーサを2つ以上選択していたらエラー)
            For intCI As Integer = 0 To intColCnt - 1
                '列ごとの選択数チェック
                intPosCnt = 0           '列全体の選択数
                intEndPosCnt = 0        'エンドブロックの選択数
                intWirPosCnt = 0        '配線ブロックの選択数
                intValPosCnt = 0        'バルブ＆マスキングプレートの選択数
                intPrtPosCnt = 0        '仕切りブロックの選択数
                intSpcPosCnt = 0        'スペーサの選択数

                strXY = CST_BLANK

                For intRI As Integer = 0 To CST_ROW - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        If intRI >= Siyou_16.EndL - 1 And _
                           intRI <= Siyou_16.EndR - 1 Then
                            'エンドブロックが選択されている場合
                            intEndPosCnt = intEndPosCnt + 1
                        ElseIf intRI = Siyou_16.Elect - 1 Then
                            '配線ブロックが選択されている場合
                            intWirPosCnt = intWirPosCnt + 1
                        ElseIf intRI >= Siyou_16.Valve1 - 1 And _
                               intRI <= Siyou_16.MPValve2 - 1 Then
                            'バルブ＆マスキングプレートが選択されている場合
                            intValPosCnt = intValPosCnt + 1
                        ElseIf intRI = Siyou_16.PartitionBlk1 - 1 Then
                            '仕切りブロックが選択されている場合
                            intPrtPosCnt = intPrtPosCnt + 1
                        ElseIf intRI >= Siyou_16.Spacer1 - 1 And _
                               intRI <= Siyou_16.RegulatorB - 1 Then
                            'スペーサが選択されている場合
                            intSpcPosCnt = intSpcPosCnt + 1
                        End If
                        intPosCnt = intPosCnt + 1
                        strXY = strXY & CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe
                    End If
                Next

                '1つの列で３個以上選択されていたらエラー
                'エンドブロックが同じ列に選択されていたらエラー
                'バルブ＆マスキングプレートが同じ列に選択されていたらエラー
                'スペーサが同じ列に選択されていたらエラー
                If intPosCnt > 2 Or intEndPosCnt > 1 Or intValPosCnt > 1 Or intSpcPosCnt > 1 Then
                    strMsg = strXY
                    strMsgCd = "W1390"
                    Exit Function
                End If

                'スペーサとバルブブロック以外の組合せはエラー
                If (intEndPosCnt > 0 And intSpcPosCnt > 0) Or _
                   (intEndPosCnt > 0 And intValPosCnt > 0) Or _
                   (intWirPosCnt > 0 And (intEndPosCnt > 0 Or intValPosCnt > 0 Or intSpcPosCnt > 0)) Or _
                   (intPrtPosCnt > 0 And (intEndPosCnt > 0 Or intValPosCnt > 0 Or intWirPosCnt > 0)) Then
                    strMsg = strXY
                    strMsgCd = "W1390"
                    Exit Function
                End If

            Next

            '---------------------------------------------
            '②ｴﾝﾄﾞﾌﾞﾛｯｸﾁｪｯｸ
            '-- 最後がｴﾝﾄﾞﾌﾞﾛｯｸ(右):2行目でない場合
            strXY = CST_BLANK
            If Not arySelectInf(Siyou_16.EndR - 1)(intRightCol) = "1" Then
                ''ｴﾗｰｾﾙ取得
                strMsg = CStr(Siyou_16.EndR) & strComma & CStr(intRightCol + 2)
                strMsgCd = "W1650"
                Exit Try
            End If

            '-- 中間行の最右列に選択がある場合
            For intR As Integer = Siyou_16.Partition1 - 1 To Siyou_16.Partition2 - 1
                If arySelectInf(intR)(intRightCol) = "1" Then
                    'ｴﾗｰｾﾙ取得
                    strXY = strXY & "0" & strComma & CStr(intRightCol + 1)
                    Exit For
                End If
            Next
            If Not strXY = CST_BLANK Then
                strMsg = strXY
                strMsgCd = "W1650"
                Exit Try
            End If

            '-- ｴﾝﾄﾞﾌﾞﾛｯｸ(右):2行目 or ｴﾝﾄﾞﾌﾞﾛｯｸ(左):1行目が2つ以上選択されている場合
            intFlg2_R = 0
            intFlg2_L = 0
            For intC As Integer = 0 To intRightCol
                If arySelectInf(Siyou_16.EndR - 1)(intC) = "1" Then
                    intFlg2_R = intFlg2_R + 1
                    strXY_R = strXY_R & Siyou_16.EndR & strComma & CStr(intC + 1) & strPipe
                End If
                If arySelectInf(Siyou_16.EndL - 1)(intC) = "1" Then
                    intFlg2_L = intFlg2_L + 1
                    strXY_L = strXY_L & Siyou_16.EndL & strComma & CStr(intC + 1) & strPipe
                End If
            Next
            If intFlg2_R > 1 Then
                'message:W1100=エンドブロックを複数指定することはできません。
                strMsg = strXY_R
                strMsgCd = "W1100"
                Exit Try
            End If
            If intFlg2_L > 1 Then
                'message:W1100=エンドブロックを複数指定することはできません。
                strMsg = strXY_L
                strMsgCd = "W1100"
                Exit Try
            End If

            '-- ｴﾝﾄﾞﾌﾞﾛｯｸ(左):1行目とｴﾝﾄﾞﾌﾞﾛｯｸ(右):2行目の形番要素の選択値がともに"X"を含む場合
            If (strKataValues(0).IndexOf("X") >= 0) And (strKataValues(1).IndexOf("X") >= 0) Then
                'message:W4050=Xを含むエンドブロックは1つしか選択できません。
                strMsgCd = "W4050"
                Exit Try
            End If

            '---------------------------------------------
            '③ｿﾚﾉｲﾄﾞ点数ﾁｪｯｸ
            Select Case strSeriesKata
                Case "MW4GB4", "MW4GZ4"
                    If strKeyKata = "S" Or strKeyKata = "Y" Then        'RM1805036_"Y"追加
                        For idx As Integer = 0 To strKataValues.Count - 1
                            Select Case idx
                                '電磁弁付ﾊﾞﾙﾌﾞﾌﾞﾛｯｸとMPV付ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ
                                Case Siyou_16.Valve1 - 1 To Siyou_16.MPValve2 - 1
                                    For intC As Integer = 0 To intColCnt - 1
                                        If arySelectInf(idx)(intC) = "1" Then
                                            If strKataValues(idx).IndexOf("MP") = -1 Then
                                                If Mid(strKataValues(idx), 7, 1) = "1" Then
                                                    If Not objKtbnStrc.strcSelection.strOpSymbol(5).ToString = "W" Then
                                                        intSoreCnt = intSoreCnt + 1
                                                    Else
                                                        intSoreCnt = intSoreCnt + 2
                                                    End If
                                                Else
                                                    intSoreCnt = intSoreCnt + 2
                                                End If
                                            Else
                                                Select Case Mid(strKataValues(idx), 10, 1)
                                                    Case "S"
                                                        intSoreCnt = intSoreCnt + 1
                                                    Case "D"
                                                        intSoreCnt = intSoreCnt + 2
                                                End Select
                                            End If
                                        End If
                                    Next
                            End Select
                        Next

                        Dim fncGetMaxSol As Integer = 16

                        Select Case objKtbnStrc.strcSelection.strOpSymbol(7).ToString
                            Case "Y10", "Y20", "Y30", "Y40"
                                fncGetMaxSol = 16
                            Case "Y11", "Y21", "Y31", "Y41"
                                fncGetMaxSol = 12
                            Case "Y12", "Y22", "Y32", "Y42"
                                fncGetMaxSol = 8
                            Case Else
                                fncGetMaxSol = 16
                        End Select

                        If intSoreCnt > fncGetMaxSol Then
                            'message:W1150=ソレノイド点数が多すぎます。
                            strMsgCd = "W1150"
                            Exit Try
                        End If
                    End If
            End Select

            '---------------------------------------------
            '④配線ﾌﾞﾛｯｸﾁｪｯｸ
            strXY = CST_BLANK
            For idx As Integer = 0 To intRightCol
                If arySelectInf(Siyou_16.Elect - 1)(idx) = "1" Then
                    strXY = strXY & Siyou_16.Elect & strComma & CStr(idx + 1) & strPipe
                    intCnt4 = intCnt4 + 1
                End If
            Next
            If intCnt4 > 1 Then
                'message:W1530=配線ブロックを複数指定することはできません。
                strMsg = strXY
                strMsgCd = "W1530"
                Exit Try
            End If

            Select Case strSeriesKata
                Case "MW4GB4", "MW4GZ4"
                    If strKeyKata = "S" Or strKeyKata = "Y" Then        'RM1805036_"Y"追加

                        If Not Right(strKataValues(4), 1) = "R" Then
                            If objKtbnStrc.strcSelection.strOpSymbol(7).ToString.Contains("Y") Then

                            Else
                                If objKtbnStrc.strcSelection.strOpSymbol(4).ToString.StartsWith("T7") Then
                                Else
                                    If Not arySelectInf(Siyou_16.Elect - 1)(1) = "1" Then
                                        'message:W4060=配線ブロックは接続位置=2に指定してください｡
                                        strMsg = Siyou_16.Elect & strComma & "2"
                                        strMsgCd = "W4060"
                                        Exit Try
                                    End If
                                End If
                            End If
                        Else
                            If Not arySelectInf(Siyou_16.Elect - 1)(intRightCol - 1) = "1" Then
                                'message:W4070=配線ブロックは右側エンドブロックの左隣に指定してください｡
                                strMsg = Siyou_16.Elect & strComma & CStr(intRightCol)
                                strMsgCd = "W4070"
                                Exit Try
                            End If
                        End If
                    End If
            End Select

            '---------------------------------------------
            '⑤電磁弁付ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ&MP付ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ使用数ﾁｪｯｸ
            For idx As Integer = Siyou_16.Valve1 - 1 To Siyou_16.MPValve2 - 1
                If Not strKataValues(idx).Trim = "" Then

                    intCnt5 = 0
                    For intC As Integer = 0 To intColCnt - 1
                        If arySelectInf(idx)(intC) = "1" Then
                            intCnt5 = intCnt5 + 1
                        End If
                    Next
                    If objKtbnStrc.strcSelection.strOpSymbol(1).ToString = "8" Then
                        Select Case Mid(strKataValues(idx), 7, 1)
                            Case "1"
                                intMixChk(1) = intCnt5
                            Case "2"
                                intMixChk(2) = intCnt5
                            Case "3"
                                intMixChk(3) = intCnt5
                            Case "4"
                                intMixChk(4) = intCnt5
                            Case "5"
                                intMixChk(5) = intCnt5
                        End Select
                        If idx >= Siyou_16.MPValve1 - 1 And idx <= Siyou_16.MPValve2 - 1 Then
                            intMixChk(6) = intMixChk(6) + intCnt5
                        End If
                    End If
                    intTtlCnt5 = intTtlCnt5 + intCnt5
                End If
            Next

            Dim strRensu As String = String.Empty
            If strSeriesKata = "MW4GB4" Or strSeriesKata = "MW4GZ4" Then
                If strKeyKata = "S" Or strKeyKata = "Y" Then        'RM1805036_"Y"追加
                    strRensu = objKtbnStrc.strcSelection.strOpSymbol(9)
                Else
                    strRensu = objKtbnStrc.strcSelection.strOpSymbol(8)
                End If
            Else
                strRensu = objKtbnStrc.strcSelection.strOpSymbol(8)
            End If

            If intTtlCnt5 > CInt(strRensu.ToString) Then
                'message:W1170=選択した電磁弁の連数が指定した値より多いです。
                strMsgCd = "W1170"
                Exit Try
            End If
            If intTtlCnt5 < CInt(strRensu.ToString) Then
                'message:W1180=選択した電磁弁の連数が指定した値より少ないです。
                strMsgCd = "W1180"
                Exit Try
            End If

            If objKtbnStrc.strcSelection.strOpSymbol(1).ToString = "8" Then
                For idx As Integer = 1 To 6
                    If intMixChk(idx) >= 1 Then
                        intFlg5 = intFlg5 + 1
                    End If
                Next
                If intFlg5 < 2 Then
                    'message:W1190=電磁弁の切換位置は２種類以上選択してください。
                    strMsgCd = "W1190"
                    Exit Try
                End If

                If objKtbnStrc.strcSelection.strOpSymbol(5).ToString.Trim = "W" Then
                    'RM1704047 MW4GB4の場合は以下の処理を行わないよう修正  2017/04/20 修正
                    'If (strSeriesKata = "MW4GB4" And strKeyKata = "S") Or (strSeriesKata = "MW4GZ4" And strKeyKata = "S") Then
                    If (strSeriesKata = "MW4GB4" And (strKeyKata = "S" Or strKeyKata = "Y")) Then       'RM1805036_"Y"追加

                    Else
                        If (strSeriesKata = "MW4GZ4" And (strKeyKata = "S" Or strKeyKata = "Y")) Then       'RM1805036_"Y"追加
                            If intMixChk(1) < 1 And intMixChk(6) < 1 Then 'シングルもMP付きどちらも選択されなかった場合
                                'message:W4080=ダブル配線(W)の時は、シングルソレノイドの選択が必要です。
                                strMsgCd = "W4080"
                                Exit Try
                            End If
                        Else
                            If intMixChk(1) < 1 Then
                                'message:W4080=ダブル配線(W)の時は、シングルソレノイドの選択が必要です。
                                strMsgCd = "W4080"
                                Exit Try
                            End If
                        End If
                    End If
                End If
            End If

            '---------------------------------------------
            '⑥複数選択ﾁｪｯｸ
            For intC As Integer = 0 To intColCnt - 1
                intCnt6_1 = 0   '1～3行目の選択されているｾﾙの数
                For intR As Integer = 0 To Siyou_16.Elect - 1
                    If arySelectInf(intR)(intC) = "1" Then
                        intCnt6_1 = intCnt6_1 + 1
                    End If
                Next
                intCnt6_2 = 0   '4～10行目の選択されているｾﾙの数
                For intR As Integer = Siyou_16.Valve1 - 1 To Siyou_16.MPValve2 - 1
                    If arySelectInf(intR)(intC) = "1" Then
                        intCnt6_2 = intCnt6_2 + 1
                    End If
                Next
                If intCnt6_1 + intCnt6_2 > 1 Then
                    'message:W1390=ブロック指定に誤りがあります。
                    strMsg = "0" & strComma & CStr(intC + 1)
                    strMsgCd = "W1390"
                    Exit Try
                End If

                intCnt6_3 = 0   '11～12行目の選択されているｾﾙの数
                For intRow As Integer = Siyou_16.Spacer1 - 1 To Siyou_16.RegulatorB - 1
                    If arySelectInf(intRow)(intC) = "1" Then
                        intCnt6_3 = intCnt6_3 + 1
                    End If
                Next
                If intCnt6_3 > 1 Then
                    'message:W1390=ブロック指定に誤りがあります。
                    strMsg = "0" & strComma & CStr(intC + 1)
                    strMsgCd = "W1390"
                    Exit Try
                End If
                If intCnt6_3 > 0 Then
                    If intCnt6_2 = 0 Then
                        'message:W1390=ブロック指定に誤りがあります。
                        strMsg = "0" & strComma & CStr(intC + 1)
                        strMsgCd = "W1390"
                        Exit Try
                    End If
                End If
            Next

            '---------------------------------------------
            Dim str() As String = Nothing
            Dim strOp_Z1 As String = String.Empty
            Dim strOp_Z3 As String = String.Empty
            Dim strOp_Z7 As String = String.Empty

            If strSeriesKata = "MW4GB4" Or strSeriesKata = "MW4GZ4" Then
                Select Case strKeyKata
                    Case "S", "Y"       'RM1805036_"Y"追加
                        str = objKtbnStrc.strcSelection.strOpSymbol(8).ToString.Split(strComma)
                    Case Else
                        str = objKtbnStrc.strcSelection.strOpSymbol(7).ToString.Split(strComma)
                End Select
            Else
                str = objKtbnStrc.strcSelection.strOpSymbol(7).ToString.Split(strComma)
            End If

            For inti As Integer = 0 To str.Length - 1
                Select Case str(inti).Trim
                    Case "Z1"
                        strOp_Z1 = "Z1"
                    Case "Z3"
                        strOp_Z3 = "Z3"
                    Case "Z7"
                        strOp_Z7 = "Z7"
                End Select
            Next

            '⑦単独給気ｽﾍﾟｰｻ/単独排気ｽﾍﾟｰｻﾁｪｯｸ
            If strOp_Z1 = "Z1" Then
                intCnt7 = 0
                For idx As Integer = 0 To intColCnt - 1
                    If arySelectInf(Siyou_16.Spacer1 - 1)(idx) = "1" Then
                        intCnt7 = intCnt7 + 1
                    End If
                Next
                If intCnt7 < 1 Then
                    'message:W4100=単独給気スペーサを1つ以上選択してください。
                    strMsgCd = "W4100"
                    Exit Try
                End If
            End If
            If strOp_Z3 = "Z3" Then
                intCnt7 = 0
                For idx As Integer = 0 To intColCnt - 1
                    If arySelectInf(Siyou_16.Spacer2 - 1)(idx) = "1" Then
                        intCnt7 = intCnt7 + 1
                    End If
                Next
                If intCnt7 < 1 Then
                    'message:W4110=単独排気スペーサを1つ以上選択してください。
                    strMsgCd = "W4110"
                    Exit Try
                End If
            End If
            If strOp_Z7 = "Z7" Then
                intCnt7 = 0
                For idx As Integer = 0 To intColCnt - 1
                    If arySelectInf(Siyou_16.RegulatorP - 1)(idx) = "1" Or _
                       arySelectInf(Siyou_16.RegulatorA - 1)(idx) = "1" Or _
                       arySelectInf(Siyou_16.RegulatorB - 1)(idx) = "1" Then
                        intCnt7 = intCnt7 + 1
                    End If
                Next
                If intCnt7 < 1 Then
                    'message:W8520=スペーサ形レギュレータを1つ以上選択してください。
                    strMsgCd = "W8520"
                    Exit Try
                End If
            End If

            '---------------------------------------------
            '⑧仕切りﾌﾟﾗｸﾞP&単独給気ｽﾍﾟｰｻ組合せﾁｪｯｸ
            For idx As Integer = 0 To intColCnt - 2
                '左から順番に選択列番号を取得し、配列に格納する(仕切りプラグ1)
                If arySelectInf(Siyou_16.Partition1 - 1)(idx) = "1" Then
                    If strFlg8_1(intCnt8_1) Is Nothing Then
                        strFlg8_1(intCnt8_1) = CStr(idx)
                    End If
                    intCnt8_1 = intCnt8_1 + 1
                End If
            Next
            If intCnt8_1 > 1 Then
                For idx_1 As Integer = 0 To intCnt8_1 - 2
                    '配列のn番目とn+1番目の間の単独給気ｽﾍﾟｰｻが選択されているかﾁｪｯｸする
                    For idx_2 As Integer = CInt(strFlg8_1(idx_1)) + 1 To CInt(strFlg8_1(idx_1 + 1))
                        If Not arySelectInf(Siyou_16.Spacer1 - 1)(idx_2) = "1" Then
                            'message:W4120=仕切プラグPとPの間に単独給気スペーサを指定してください。
                            strXY = Siyou_16.Partition1 & strComma & CStr(CInt(strFlg8_1(idx_1)) + 1) & strPipe & Siyou_16.Partition1 & strComma & CStr(CInt(strFlg8_1(idx_1 + 1)) + 1)
                            strMsg = strXY
                            strMsgCd = "W4120"
                            Exit Try
                        End If
                    Next
                Next
            End If

            '---------------------------------------------
            '⑨仕切りﾌﾟﾗｸﾞR&単独排気ｽﾍﾟｰｻ組合せﾁｪｯｸ　
            For idx As Integer = 0 To intColCnt - 2
                '左から順番に選択列番号を取得する(14行目)
                If arySelectInf(Siyou_16.Partition2 - 1)(idx) = "1" Then
                    If strFlg9_1(intCnt9_1) = "" Then
                        strFlg9_1(intCnt9_1) = CStr(idx)
                    End If
                    intCnt9_1 = intCnt9_1 + 1
                End If
            Next
            If intCnt9_1 > 1 Then
                For idx_1 As Integer = 0 To intCnt9_1 - 2
                    '配列のn番目とn+1番目の間の単独排気ｽﾍﾟｰｻが選択されているかﾁｪｯｸする
                    For idx_2 As Integer = CInt(strFlg9_1(idx_1)) + 1 To CInt(strFlg9_1(idx_1 + 1))
                        If Not arySelectInf(Siyou_16.Spacer2 - 1)(idx_2) = "1" Then
                            'message:W4150=仕切プラグRとRの間に単独給気スペーサを指定してください。
                            strXY = Siyou_16.Partition2 & strComma & CStr(CInt(strFlg9_1(idx_1)) + 1) & strPipe & Siyou_16.Partition2 & strComma & CStr(CInt(strFlg9_1(idx_1 + 1)) + 1)
                            strMsg = strXY
                            strMsgCd = "W4150"
                            Exit Try
                        End If
                    Next
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
    '*   fncInputChk②
    '*【処理】
    '*   入力内容をチェックする
    '********************************************************************************************
    Public Shared Function fncInpCheck2(objKtbnStrc As KHKtbnStrc, ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        fncInpCheck2 = False
        Try
            '---------------------------------------------
            '品名ﾘｽﾄｺﾝﾄﾛｰﾙ　数値ﾃｷｽﾄ入力値ﾁｪｯｸ
            '- 15～19行目
            If Not SiyouBLL.fncOtherKataCheck(objKtbnStrc, Siyou_16.Plug1, Siyou_16.Inspect, 0, strMsgCd) Then
                Exit Function
            End If

            '---------------------------------------------
            '形番選択ﾁｪｯｸ
            For idx As Integer = 0 To intPosRowCnt - 1
                If CInt(objKtbnStrc.strcSelection.intQuantity(idx)) > 0 Then
                    If objKtbnStrc.strcSelection.strOptionKataban(idx).Trim = CST_BLANK Then
                        'message:W1400=形番を選択してください。
                        strMsgCd = "W1400"
                        Exit Try
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try

    End Function
End Class
