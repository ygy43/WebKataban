Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_15

    Public Shared intPosRowCnt As Integer = 30       'RM1803032_スペーサ行追加
    Public Shared intColCnt As Integer = 40          'RM1803032_マニホールド連数拡張

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
    '*   fncInpCheck1
    '*【処理】
    '*   入力チェック１
    '********************************************************************************************
    Public Shared Function fncInpCheck1(objKtbnStrc As KHKtbnStrc, ByRef strMsg As String, ByRef strMsgCd As String) As Boolean

        Dim intCnt As Integer
        Dim intRightEdge As Integer
        Dim intLeftEdge As Integer
        Dim intLoop As Integer
        Dim intNum As Integer
        Dim intPartREdge As Integer
        Dim intMixSwitch() As Integer = Nothing
        Dim bolFlag1 As Boolean
        Dim bolFlag2 As Boolean
        Dim bolFlag3 As Boolean
        Dim bolFlag4 As Boolean
        Dim bolPart As Boolean
        Dim bolExht As Boolean
        Dim bolH As Boolean
        Dim sbCoordinates As New System.Text.StringBuilder
        Dim strCoordinates As String

        fncInpCheck1 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            '******* 未接続位置チェック(7.1) ***************************************
            For intCI As Integer = 0 To intColCnt - 1
                intLoop = 0
                Do While intLoop < arySelectInf.Count - 1
                    If arySelectInf(intLoop)(intCI) = "1" Then
                        '一番左側の選択列Ｎｏを取得
                        intLeftEdge = intCI + 1
                        Exit For
                    End If
                    intLoop = intLoop + 1
                Loop
            Next
            '選択セルが一つもない場合、エラー
            If intLeftEdge = 0 Then
                strMsgCd = "W1030"
                Exit Try
            End If

            For intCI As Integer = intColCnt - 1 To intLeftEdge Step -1
                bolFlag1 = False
                For intRI As Integer = 0 To arySelectInf.Count - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        '一番右側の選択列Ｎｏを取得
                        If intRightEdge = 0 Then
                            intRightEdge = intCI + 1
                        End If

                        bolFlag1 = True
                        Exit For
                    End If
                Next
                '中間に一つも選択されていない列がある場合、エラー
                If intRightEdge > 0 And Not bolFlag1 Then
                    strCoordinates = "0" & strComma & CStr(intCI + 1)
                    strMsg = strCoordinates
                    strMsgCd = "W1020"
                    Exit Try
                End If
            Next

            '******* 入出力ブロックチェック(7.2 / 7.3) ******************************
            If Int(strUseValues(Siyou_15.InOut1 - 1)) > 0 And _
               Len(Trim(strKataValues(Siyou_15.InOut1 - 1))) = 0 Then
                strCoordinates = CStr(Siyou_15.InOut1) & strComma & "0"
                strMsg = strCoordinates
                strMsgCd = "W1600"
                Exit Try
            End If
            If Int(strUseValues(Siyou_15.InOut2 - 1)) > 0 And _
               Len(Trim(strKataValues(Siyou_15.InOut2 - 1))) = 0 Then
                strCoordinates = CStr(Siyou_15.InOut2) & strComma & "0"
                strMsg = strCoordinates
                strMsgCd = "W1610"
                Exit Try
            End If

            '******* エンドブロックチェック(7.4) ************************************
            If Int(strUseValues(Siyou_15.EndL - 1)) > 0 And _
               Len(Trim(strKataValues(Siyou_15.EndL - 1))) = 0 Then
                strCoordinates = CStr(Siyou_15.EndL) & strComma & "0"
                strMsg = strCoordinates
                strMsgCd = "W1620"
                Exit Try
            End If

            For intRI As Integer = Siyou_15.EndR - 1 To Siyou_15.EndL - 1
                If Int(strUseValues(intRI)) > 1 Then
                    For intCI As Integer = intLeftEdge - 1 To intRightEdge - 1
                        If arySelectInf(intRI)(intCI) = "1" Then
                            sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                        End If
                    Next
                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                    strMsg = strCoordinates
                    strMsgCd = "W1100"
                    Exit Try
                End If
            Next

            '一番右側の選択列でエンドブロック(右)以外が選択されている場合、エラー
            bolFlag1 = True
            For intRI As Integer = 0 To arySelectInf.Count - 1
                If intRightEdge > 0 Then
                    If arySelectInf(intRI)(intRightEdge - 1) = "1" Then
                        If intRI <> Siyou_15.EndR - 1 Then
                            bolFlag1 = False
                            Exit For
                        End If
                    End If
                End If
            Next
            If Not bolFlag1 Then
                sbCoordinates.Append(CStr(Siyou_15.EndR) & strComma & (intRightEdge + 1) & strPipe)
                strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                strMsg = strCoordinates
                strMsgCd = "W1650"
                Exit Try
            End If
            sbCoordinates = New System.Text.StringBuilder

            '******* 電装ブロックチェック(7.5) *************************************
            If Int(strUseValues(Siyou_15.Elect - 1)) > 1 Then
                strCoordinates = CStr(Siyou_15.Elect) & strComma & "0"
                strMsg = strCoordinates
                strMsgCd = "W1640"
                Exit Try
            End If

            '電装ブロック選択列より右に入出力ブロック選択列が存在する場合、エラー
            intNum = intColCnt
            For intCI As Integer = intLeftEdge - 1 To intRightEdge - 1
                If arySelectInf(Siyou_15.Elect - 1)(intCI) = "1" Then
                    intNum = intCI + 1
                    Exit For
                End If
            Next
            For intCI As Integer = intNum To intRightEdge - 1
                For intRI As Integer = Siyou_15.InOut1 - 1 To Siyou_15.InOut2 - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        strCoordinates = CStr(intRI + 1) & strComma & "0"
                        strMsg = strCoordinates
                        strMsgCd = "W1630"
                        Exit Try
                    End If
                Next
            Next

            '******* ソレノイド点数チェック(7.6) *********************************
            intCnt = fncGetSolenoidCnt(objKtbnStrc)
            If intCnt > KHKataban.fncGetMaxSol(objKtbnStrc.strcSelection.strOpSymbol, 15) Then
                strMsgCd = "W1150"
                Exit Try
            End If

            '******* バルブブロック/接続口径の使用数チェック(7.7) *****************
            subGetMixCheckVal(objKtbnStrc, intMixSwitch, intNum)
            If intNum > Int(objKtbnStrc.strcSelection.strOpSymbol(8).ToString) Then
                strMsgCd = "W1170"
                Exit Try

            ElseIf intNum < Int(objKtbnStrc.strcSelection.strOpSymbol(8).ToString) Then
                strMsgCd = "W1180"
                Exit Try
            End If

            'ミックスチェック値(切替位置区分)の確認
            intCnt = 0
            bolH = True
            If Left(objKtbnStrc.strcSelection.strOpSymbol(1).ToString, 1) = "8" Then
                For intI As Integer = 0 To UBound(intMixSwitch)
                    If intMixSwitch(intI) > 0 Then
                        intCnt = intCnt + 1
                    End If
                Next
                If intCnt < 2 Then
                    strMsgCd = "W1190"
                    Exit Try
                End If

                Dim strOptionH As String = Nothing
                Dim str_6() As String = objKtbnStrc.strcSelection.strOpSymbol(6).ToString.Split(strComma)
                For intI As Integer = 0 To str_6.Length - 1
                    If str_6(intI).Contains("H") Then
                        strOptionH = "H"
                    End If
                Next
                If strOptionH IsNot Nothing Then
                    If intMixSwitch(0) = 0 And intMixSwitch(1) = 0 And intMixSwitch(2) = 0 And _
                       intMixSwitch(3) = 0 And intMixSwitch(4) > 0 And intMixSwitch(5) = 0 And _
                       intMixSwitch(6) > 0 And intMixSwitch(7) = 0 Then
                        bolH = False
                    ElseIf intMixSwitch(0) = 0 And intMixSwitch(1) = 0 And intMixSwitch(2) = 0 And _
                       intMixSwitch(3) = 0 And intMixSwitch(4) > 0 And intMixSwitch(5) = 0 And _
                       intMixSwitch(6) = 0 And intMixSwitch(7) > 0 Then
                        bolH = False
                    ElseIf intMixSwitch(0) = 0 And intMixSwitch(1) = 0 And intMixSwitch(2) = 0 And _
                       intMixSwitch(3) = 0 And intMixSwitch(4) = 0 And intMixSwitch(5) = 0 And _
                       intMixSwitch(6) > 0 And intMixSwitch(7) > 0 Then
                        bolH = False
                    ElseIf intMixSwitch(0) = 0 And intMixSwitch(1) = 0 And intMixSwitch(2) = 0 And _
                       intMixSwitch(3) = 0 And intMixSwitch(4) > 0 And intMixSwitch(5) = 0 And _
                       intMixSwitch(6) > 0 And intMixSwitch(7) > 0 Then
                        bolH = False
                    ElseIf intMixSwitch(0) = 0 And intMixSwitch(1) = 0 And intMixSwitch(2) = 0 And _
                       intMixSwitch(3) = 0 And intMixSwitch(4) > 0 And intMixSwitch(5) = 0 And _
                       intMixSwitch(6) > 0 And intMixSwitch(7) > 0 Then
                        bolH = False
                    End If
                    If Not bolH Then
                        strMsgCd = "W1660"
                        Exit Try
                    End If
                End If
            End If

            'ミックスチェック値(接続口径)の確認
            intCnt = 0
            If Left(objKtbnStrc.strcSelection.strOpSymbol(3).ToString, 2) = "CX" Then
                If Not SiyouBLL.fncMixBlockCheck(objKtbnStrc, Siyou_15.Valve1 - 1, Siyou_15.Valve8 - 1, strMsgCd) Then
                    Exit Function
                End If
            End If

            '******* 電磁弁付バルブブロック継手チェック(7.8) **************
            If Not SiyouBLL.fncBlockCheck(strUseValues, strKataValues, Siyou_15.Valve1 - 1, Siyou_15.Valve8 - 1, strMsgCd) Then
                Exit Function
            End If
            sbCoordinates = New System.Text.StringBuilder

            '******* 給排気ブロック単独チェック(7.9) ****
            bolFlag1 = False
            bolFlag2 = False
            bolFlag3 = False
            For intRI As Integer = Siyou_15.Exhaust1 - 1 To Siyou_15.Exhaust2 - 1
                If Int(strUseValues(intRI)) > 0 Then
                    If strKataValues(intRI).Contains("-QZ") Then
                        bolFlag1 = True
                    End If
                    If ClsInputCheck_01.fncContaints(strKataValues(intRI), "-Q-,-QK-,-QKZ-") Then
                        bolFlag2 = True
                    End If
                    If strKataValues(intRI).Contains("-SA") Then
                        bolFlag3 = True
                    End If
                    For intCI As Integer = intLeftEdge - 1 To intRightEdge - 1
                        If arySelectInf(intRI)(intCI) = "1" Then
                            sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                        End If
                    Next
                End If
            Next
            '給排気ブロックの選択型番に"-QZ"が含まれている場合
            If bolFlag1 Then
                '給排気ブロックの選択型番に"-Q-"も″"-QK-"も"-QKZ-"も含まれていない場合
                If Not bolFlag2 Then
                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                    strMsg = strCoordinates
                    strMsgCd = "W1680"
                    Exit Try
                End If
                '給排気ブロックの選択型番に"-SA"が含まれている場合
                If bolFlag3 Then
                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                    strMsg = strCoordinates
                    strMsgCd = "W1690"
                    Exit Try
                End If
            End If

            '******* 給排気ブロック/仕切りブロック組み合わせチェック(7.10) ****
            bolFlag1 = False
            bolFlag2 = False
            For intRI As Integer = Siyou_15.Exhaust1 - 1 To Siyou_15.Exhaust2 - 1
                If Int(strUseValues(intRI)) > 0 And _
                   strKataValues(intRI).Contains("-QZ-") Then

                    bolFlag1 = True
                    For intCI As Integer = intLeftEdge - 1 To intRightEdge - 1
                        If arySelectInf(intRI)(intCI) = "1" Then
                            sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                        End If
                    Next
                End If
            Next
            For intRI As Integer = Siyou_15.Partition1 - 1 To Siyou_15.Partition2 - 1
                If Int(strUseValues(intRI)) > 0 And _
                   strKataValues(intRI).Contains("-SA") Then

                    bolFlag2 = True
                End If
            Next
            '給排気ブロックの選択型番に"-QZ"が含まれており、仕切りブロックの選択型番に"-SA"が含まれている場合
            If bolFlag1 And bolFlag2 Then
                strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                strMsgCd = "W2740"
                Exit Try
            End If
            sbCoordinates = New System.Text.StringBuilder

            '******* 仕切りブロック/給排気ブロック/バルブブロック/エンドブロック組み合わせチェック(7.11) ******
            bolPart = False
            bolExht = False
            For intRI As Integer = Siyou_15.Partition1 - 1 To Siyou_15.Partition2 - 1
                '仕切りブロックに選択がある場合
                If Int(strUseValues(intRI)) > 0 Then
                    bolPart = True

                    '全ての選択列に対して以下の処理を行う
                    For intCI As Integer = intLeftEdge - 1 To intRightEdge - 1
                        If arySelectInf(intRI)(intCI) = "1" Then

                            '仕切りブロック選択列より右にバルブブロック選択列が存在するか確認(7.11.1)
                            If Not fncFoundValve(objKtbnStrc, 1, intCI + 1, intRightEdge) Then
                                strMsgCd = "W1710"
                                Exit Try
                            End If

                            'Del by Zxjike 2013/10/21
                            '仕切りブロック選択列より左に給排気ブロック選択列が存在するか確認(7.11.4)
                            'If Not fncFoundExhaust(2, intCI - 1, intLeftEdge, arySelectInf, strKataValues) Then
                            '    strMsgCd = "W2270"
                            '    Exit Try
                            'End If
                            'Add by Zxjike 2013/10/21
                            '仕切りブロック選択列より左に給排気ブロック選択列が存在するか確認
                            If Not fncFoundExhaust(objKtbnStrc, 0, intCI - 1, intLeftEdge - 1) Then
                                strMsgCd = "W2270"
                                Exit Try
                            End If

                            '******* 仕切りブロック/給排気ブロック組み合わせチェック(7.12) ******
                            '仕切りブロック選択列の左に、形番に"X"を含む給排気ブロック選択列があるか確認
                            If fncCheckExhaust(objKtbnStrc, 2, intCI - 1, intLeftEdge - 1) Then
                                strMsgCd = "W1690"
                                Exit Try
                            End If

                        End If
                    Next

                End If
            Next

            If bolPart Then
                '一番左の選択列に対して以下の処理を行う
                intLoop = intLeftEdge - 1
                Do While intLoop < intRightEdge
                    If arySelectInf(Siyou_15.Partition1 - 1)(intLoop) = "1" Or _
                       arySelectInf(Siyou_15.Partition2 - 1)(intLoop) = "1" Then

                        '仕切りブロック選択列より左にバルブブロック選択列が存在するか確認(7.11.1)
                        If Not fncFoundValve(objKtbnStrc, 0, intLoop - 1, intLeftEdge) Then
                            strMsgCd = "W1710"
                            Exit Try
                        End If

                        Exit Do
                    End If
                    intLoop = intLoop + 1
                Loop

                '一番右の選択列に対して以下の処理を行う
                intLoop = intRightEdge - 1
                Do While intLoop >= intLeftEdge - 1
                    If arySelectInf(Siyou_15.Partition1 - 1)(intLoop) = "1" Or _
                       arySelectInf(Siyou_15.Partition2 - 1)(intLoop) = "1" Then

                        '仕切りブロック選択列より右に給排気ブロック選択列が存在するか確認(7.11.4)
                        If Not fncFoundExhaust(objKtbnStrc, 1, intLoop + 1, intRightEdge) Then
                            If Not (Int(strUseValues(Siyou_15.EndR - 1)) > 0 And _
                                    strKataValues(Siyou_15.EndR - 1).Contains("X")) Then
                                strMsgCd = "W2270"
                                Exit Try
                            End If
                        End If

                        intPartREdge = intLoop
                        Exit Do
                    End If
                    intLoop = intLoop - 1
                Loop

            End If

            For intRI As Integer = Siyou_15.Exhaust1 - 1 To Siyou_15.Exhaust2 - 1
                If Int(strUseValues(intRI)) > 0 Then
                    bolExht = True
                    '給排気ブロック(仕切りタイプ)に選択がある場合
                    If strKataValues(intRI).Contains("-S") Then
                        '全ての選択列に対して以下の処理を行う
                        For intCI As Integer = intLeftEdge - 1 To intRightEdge - 1
                            If arySelectInf(intRI)(intCI) = "1" Then
                                '******* 仕切りブロック/給排気ブロック組み合わせチェック(7.13) ******
                                '右 → 左
                                If fncCheckExhaust(objKtbnStrc, 0, intCI - 1, intLeftEdge) Then
                                    strMsgCd = "W1690"
                                    Exit Try
                                End If
                                '左 → 右
                                If fncCheckExhaust(objKtbnStrc, 1, intCI + 1, intRightEdge) Then
                                    strMsgCd = "W1690"
                                    Exit Try
                                End If
                            End If
                        Next
                    End If
                End If
            Next

            If bolExht Then
                '一番左の選択列に対して以下の処理を行う
                intLoop = intLeftEdge - 1
                Do While intLoop < intRightEdge
                    If (arySelectInf(Siyou_15.Exhaust1 - 1)(intLoop) = "1" And strKataValues(Siyou_15.Exhaust1 - 1).Contains("-S")) Or _
                       (arySelectInf(Siyou_15.Exhaust2 - 1)(intLoop) = "1" And strKataValues(Siyou_15.Exhaust2 - 1).Contains("-S")) Then
                        '給排気ブロック(仕切りタイプ)選択列より左にバルブブロック選択列が存在するか確認(7.11.2)
                        If Not fncFoundValve(objKtbnStrc, 0, intLoop - 1, intLeftEdge) Then
                            strMsgCd = "W1720"
                            Exit Try
                        End If
                        '給排気ブロック(仕切りタイプ)選択列より左に給排気ブロック選択列が存在するか確認(7.11.3)
                        If Not fncFoundExhaust(objKtbnStrc, 0, intLoop - 1, intLeftEdge) Then
                            strMsgCd = "W1730"
                            Exit Try
                        End If
                        Exit Do
                    End If
                    intLoop = intLoop + 1
                Loop

                '一番右の選択列に対して以下の処理を行う
                intLoop = intRightEdge - 1
                Do While intLoop >= intLeftEdge - 1
                    If (arySelectInf(Siyou_15.Exhaust1 - 1)(intLoop) = "1" And strKataValues(Siyou_15.Exhaust1 - 1).Contains("-S")) Or _
                       (arySelectInf(Siyou_15.Exhaust2 - 1)(intLoop) = "1" And strKataValues(Siyou_15.Exhaust2 - 1).Contains("-S")) Then
                        '給排気ブロック(仕切りタイプ)選択列より右にバルブブロック選択列が存在するか確認(7.11.2)
                        If Not fncFoundValve(objKtbnStrc, 2, intLoop + 1, intRightEdge) Then
                            strMsgCd = "W1720"
                            Exit Try
                        End If
                        Exit Do
                    End If
                    intLoop = intLoop - 1
                Loop
            End If

            '仕切りブロック、給排気ブロック共に未選択の場合(7.11)
            If Not bolPart And Not bolExht Then
                strMsgCd = "W1700"
                Exit Try
            End If

            '201501月次更新
            '7.11 仕切ブロック・給排気ブロック・エンドブロック組合せチェック
            '給排気に'X'が含まれる時、エンドブロックに"X"が含まれない場合、エラー
            If (CInt(strUseValues(Siyou_15.Exhaust1 - 1)) > 0 And InStr(strKataValues(Siyou_15.Exhaust1 - 1), "X") = 0) Or _
               (CInt(strUseValues(Siyou_15.Exhaust2 - 1)) > 0 And InStr(strKataValues(Siyou_15.Exhaust2 - 1), "X") = 0) Or _
               (CInt(strUseValues(Siyou_15.Exhaust1 - 1)) = 0 And CInt(strUseValues(Siyou_15.Exhaust2 - 1)) = 0) Then
            Else
                If (CInt(strUseValues(Siyou_15.EndL - 1)) > 0 And InStr(strKataValues(Siyou_15.EndL - 1), "X") > 0) Or _
                   (CInt(strUseValues(Siyou_15.EndR - 1)) > 0 And InStr(strKataValues(Siyou_15.EndR - 1), "X") > 0) Then
                Else
                    strMsgCd = "W2260"
                    Exit Function
                End If
            End If

            '******* スペーサ/バルブブロック組み合わせチェック(7.14) ***************
            For intCI As Integer = intLeftEdge - 1 To intRightEdge - 1
                If arySelectInf(Siyou_15.Spacer1 - 1)(intCI) = "1" Or _
                   arySelectInf(Siyou_15.Spacer2 - 1)(intCI) = "1" Or _
                   arySelectInf(Siyou_15.Spacer3 - 1)(intCI) = "1" Or _
                   arySelectInf(Siyou_15.Spacer4 - 1)(intCI) = "1" Then

                    bolFlag1 = False
                    For intRI As Integer = Siyou_15.Valve1 - 1 To Siyou_15.Valve8 - 1
                        If arySelectInf(intRI)(intCI) = "1" Then
                            bolFlag1 = True
                        End If
                    Next

                    'スペーサが選択されている列でバルブブロックが一つも選択されていない場合
                    If Not bolFlag1 Then
                        If arySelectInf(Siyou_15.Spacer1 - 1)(intCI) = "1" Then
                            sbCoordinates.Append(Siyou_15.Spacer1 & strComma & CStr(intCI + 1) & strPipe)
                        End If
                        If arySelectInf(Siyou_15.Spacer2 - 1)(intCI) = "1" Then
                            sbCoordinates.Append(Siyou_15.Spacer2 & strComma & CStr(intCI + 1) & strPipe)
                        End If
                        If arySelectInf(Siyou_15.Spacer3 - 1)(intCI) = "1" Then
                            sbCoordinates.Append(Siyou_15.Spacer3 & strComma & CStr(intCI + 1) & strPipe)
                        End If
                        If arySelectInf(Siyou_15.Spacer4 - 1)(intCI) = "1" Then
                            sbCoordinates.Append(Siyou_15.Spacer4 & strComma & CStr(intCI + 1) & strPipe)
                        End If
                        strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                        strMsg = strCoordinates
                        strMsgCd = "W2250"
                        Exit Try
                    End If
                End If
            Next

            '******* スペーサ/バルブブロック組み合わせチェック２(7.15) ***************
            For intCI As Integer = intLeftEdge - 1 To intRightEdge - 1
                bolFlag1 = False
                For intRI As Integer = Siyou_15.Valve1 - 1 To Siyou_15.Valve8 - 1
                    If arySelectInf(intRI)(intCI) = "1" And _
                       strKataValues(intRI).Contains("-MP") Then
                        bolFlag1 = True
                    End If
                Next

                If bolFlag1 Then
                    For intRI As Integer = Siyou_15.Spacer1 - 1 To Siyou_15.Spacer4 - 1
                        'マスキングプレートとスペーサが同一列で選択されている場合
                        If arySelectInf(intRI)(intCI) = "1" Then
                            strCoordinates = CStr(intRI + 1) & strComma & CStr(intCI + 1)
                            strMsg = strCoordinates
                            strMsgCd = "W2300"
                            Exit Try
                        End If
                    Next
                End If
            Next

            '******* スペーサチェック(7.16) ***************
            bolFlag1 = False
            bolFlag2 = False
            bolFlag3 = False
            bolFlag4 = False
            For intRI As Integer = Siyou_15.Spacer1 - 1 To Siyou_15.Spacer4 - 1
                If Int(strUseValues(intRI)) > 0 Then
                    If strKataValues(intRI).Contains("-P") And _
                       Not strKataValues(intRI).Contains("-PC") And _
                       Not strKataValues(intRI).Contains("-PIS") Then
                        bolFlag1 = True
                    End If
                    If strKataValues(intRI).Contains("-R") Then
                        bolFlag2 = True
                    End If
                    If strKataValues(intRI).Contains("-PC") Then
                        bolFlag3 = True
                    End If
                    If strKataValues(intRI).Contains("-PIS") Then
                        bolFlag4 = True
                    End If
                End If
            Next

            Dim strOptionZ1 As String = Nothing
            Dim strOptionZ3 As String = Nothing
            Dim strOptionZ6 As String = Nothing
            Dim strOptionZ8 As String = Nothing
            Dim str() As String = objKtbnStrc.strcSelection.strOpSymbol(6).ToString.Split(strComma)
            For intI As Integer = 0 To str.Length - 1
                If str(intI).Contains("Z1") Then
                    strOptionZ1 = "Z1"
                End If
                If str(intI).Contains("Z3") Then
                    strOptionZ3 = "Z3"
                End If
                If str(intI).Contains("Z6") Then
                    strOptionZ6 = "Z6"
                End If
                If str(intI).Contains("Z8") Then
                    strOptionZ8 = "Z8"
                End If
            Next
            If strOptionZ1 IsNot Nothing And Not bolFlag1 Then
                strMsgCd = "W4030"
                Exit Try
            End If
            If strOptionZ3 IsNot Nothing And Not bolFlag2 Then
                strMsgCd = "W4040"
                Exit Try
            End If
            If strOptionZ6 IsNot Nothing And Not bolFlag3 Then
                strMsgCd = "W8500"
                Exit Try
            End If
            If strOptionZ8 IsNot Nothing And Not bolFlag4 Then
                strMsgCd = "W8510"
                Exit Try
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
    '*   入力チェック２
    '********************************************************************************************
    Public Shared Function fncInpCheck2(objKtbnStrc As KHKtbnStrc, ByRef dblStdNum As Double, _
                                  ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        Dim hshtKataban As New Hashtable
        Dim strCoordinats As String
        Dim sbBlockCoord As New System.Text.StringBuilder
        Dim sbCoordinats As New System.Text.StringBuilder
        Dim intPosCnt As Integer
        Dim intInOutPosCnt As Integer
        Dim intValPosCnt As Integer
        Dim intSpacPosCnt As Integer
        Dim intExPosCnt As Integer
        Dim intPartPosCnt As Integer
        Dim intEndPosCnt As Integer
        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban

        fncInpCheck2 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            '設置位置が選択されている行の形番が未選択の場合、エラー
            For intRI As Integer = 3 To strUseValues.Count - 1
                If Int(strUseValues(intRI - 1)) > 0 And Len(Trim(strKataValues(intRI - 1))) = 0 And _
                    intRI <> Siyou_15.Rail Then
                    strMsgCd = "W1400"
                    fncInpCheck2 = False
                    Exit Function
                End If
            Next

            '設置位置選択チェック
            For intCI As Integer = 0 To intColCnt - 1

                '列ごとの選択数チェック
                intPosCnt = 0           '列全体の選択数
                intInOutPosCnt = 0      '入出力ブロックの選択数
                intValPosCnt = 0        'バルブブロックの選択数
                intSpacPosCnt = 0       'スペーサの選択数
                intExPosCnt = 0         '給排気ブロックの選択数
                intPartPosCnt = 0       '仕切りプラグの選択数
                intEndPosCnt = 0        'エンドブロックの選択数

                sbCoordinats = Nothing
                sbCoordinats = New System.Text.StringBuilder

                For intRI As Integer = 0 To arySelectInf.Count - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        If intRI >= Siyou_15.InOut1 - 1 And intRI <= Siyou_15.InOut2 - 1 Then
                            '入出力ブロックが選択されている場合
                            intInOutPosCnt = intInOutPosCnt + 1
                        ElseIf intRI >= Siyou_15.Valve1 - 1 And intRI <= Siyou_15.Valve8 - 1 Then
                            'バルブ＆マスキングプレートが選択されている場合
                            intValPosCnt = intValPosCnt + 1
                        ElseIf intRI >= Siyou_15.Spacer1 - 1 And intRI <= Siyou_15.Spacer4 - 1 Then
                            'スペーサが選択されている場合
                            intSpacPosCnt = intSpacPosCnt + 1
                        ElseIf intRI >= Siyou_15.Exhaust1 - 1 And intRI <= Siyou_15.Exhaust2 - 1 Then
                            '給排気ブロックが選択されている場合
                            intExPosCnt = intExPosCnt + 1
                        ElseIf intRI >= Siyou_15.Partition1 - 1 And intRI <= Siyou_15.Partition2 - 1 Then
                            '仕切りプラグが選択されている場合
                            intPartPosCnt = intPartPosCnt + 1
                        ElseIf intRI >= Siyou_15.EndR - 1 And intRI <= Siyou_15.EndL - 1 Then
                            '仕切りプラグが選択されている場合
                            intEndPosCnt = intEndPosCnt + 1
                        End If
                        intPosCnt = intPosCnt + 1
                        sbCoordinats.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                    End If
                Next

                '1つの列で３個以上選択されていたらエラー
                '入出力ブロック、バルブ＆マスキングプレート、スペーサ、給排気ブロック、仕切りプラグ、エンドブロックが同列に2つ以上選択されていたらエラー
                If intPosCnt > 2 Or intInOutPosCnt > 1 Or intSpacPosCnt > 1 Or intValPosCnt > 1 Or intExPosCnt > 1 Or intPartPosCnt > 1 Or intEndPosCnt > 1 Then
                    strCoordinats = Left(sbCoordinats.ToString, Len(sbCoordinats.ToString) - 1)
                    strMsg = strCoordinats
                    strMsgCd = "W1390"
                    fncInpCheck2 = False
                    Exit Function
                End If
                '同列において、スペーサとそれ以外という組合せ以外はエラー
                If intPosCnt > 1 And intSpacPosCnt < 1 Then
                    strCoordinats = Left(sbCoordinats.ToString, Len(sbCoordinats.ToString) - 1)
                    strMsg = strCoordinats
                    strMsgCd = "W1390"
                    fncInpCheck2 = False
                    Exit Function
                End If
            Next

            '品名リストコントロール 形番リスト重複チェック
            For intRI As Integer = Siyou_15.Valve1 - 1 To Siyou_15.Valve8 - 1
                If Len(Trim(strKataValues(intRI))) = 0 Then
                ElseIf hshtKataban.ContainsKey(strKataValues(intRI)) Then
                    strMsgCd = "W1330"
                    Exit Try
                Else
                    hshtKataban.Add(strKataValues(intRI), "")
                End If
            Next

            hshtKataban = New Hashtable
            For intRI As Integer = Siyou_15.Spacer1 - 1 To Siyou_15.Spacer4 - 1
                If Len(Trim(strKataValues(intRI))) = 0 Then
                ElseIf hshtKataban.ContainsKey(strKataValues(intRI)) Then
                    strMsgCd = "W1330"
                    Exit Try
                Else
                    hshtKataban.Add(strKataValues(intRI), "")
                End If
            Next

            hshtKataban = New Hashtable
            For intRI As Integer = Siyou_15.Exhaust1 - 1 To Siyou_15.Exhaust2 - 1
                If Len(Trim(strKataValues(intRI))) = 0 Then
                ElseIf hshtKataban.ContainsKey(strKataValues(intRI)) Then
                    strMsgCd = "W1330"
                    Exit Try
                Else
                    hshtKataban.Add(strKataValues(intRI), "")
                End If
            Next

            '品名リストコントロール 使用数チェック
            If Not SiyouBLL.fncOtherKataCheck(objKtbnStrc, Siyou_15.Plug1, Siyou_15.Inspect2, _
                                     Siyou_15.Rail, strMsgCd) Then
                Exit Function
            End If

            '設定値判定
            If strKataValues(Siyou_15.Rail - 1).ToString.Length <= 0 Then strKataValues(Siyou_15.Rail - 1) = 0
            If Not SiyouBLL.fncRailchk(strKataValues(Siyou_15.Rail - 1), strUseValues(Siyou_15.Rail - 1), dblStdNum, strMsgCd) Then
                strMsg = Siyou_15.Rail & ",0"
                Exit Function
            End If

            fncInpCheck2 = True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncGetSolenoidCnt
    '*【処理】
    '*   ソレノイド点数カウント
    '********************************************************************************************
    Public Shared Function fncGetSolenoidCnt(objKtbnStrc As KHKtbnStrc) As Integer
        Dim intSoleCnt As Integer = 0
        Try
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            For intRI As Integer = Siyou_15.Valve1 - 1 To Siyou_15.Valve8 - 1
                For intCI As Integer = 0 To intColCnt - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        '対象行の選択形番にMPが含まれている場合
                        If strKataValues(intRI).Contains("MP") Then
                            If Len(strKataValues(intRI)) < 10 Then
                            ElseIf strKataValues(intRI).Substring(9, 1) = "S" Then
                                intSoleCnt = intSoleCnt + 1
                            ElseIf strKataValues(intRI).Substring(9, 1) = "D" Then
                                intSoleCnt = intSoleCnt + 2
                            End If

                            '対象行の選択型番の先頭７文字目が1の場合
                        ElseIf Len(strKataValues(intRI)) < 7 Then
                        ElseIf strKataValues(intRI).Substring(6, 1) = "1" Then

                            If objKtbnStrc.strcSelection.strOpSymbol(5).ToString = "W" Then
                                intSoleCnt = intSoleCnt + 2
                            Else
                                intSoleCnt = intSoleCnt + 1
                            End If
                        Else
                            intSoleCnt = intSoleCnt + 2
                        End If
                    End If
                Next
            Next
            fncGetSolenoidCnt = intSoleCnt
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   subGetMixCheckVal
    '*【処理】
    '*   ミックスチェック値取得
    '********************************************************************************************
    Public Shared Sub subGetMixCheckVal(objKtbnStrc As KHKtbnStrc, ByRef intMixSwitch As Integer(), _
                                        ByRef intElectCnt As Integer)

        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban

            ReDim intMixSwitch(8)
            intElectCnt = 0

            For intRI As Integer = Siyou_15.Valve1 - 1 To Siyou_15.Valve8 - 1
                If Left(objKtbnStrc.strcSelection.strOpSymbol(1).ToString, 1) = "8" Then
                    If Len(Trim(strKataValues(intRI))) > 7 Then
                        If Left(strKataValues(intRI), 4) = "NW3G" Then
                            Select Case strKataValues(intRI).Substring(6, 2)
                                Case "10"
                                    intMixSwitch(0) = intMixSwitch(0) + Int(strUseValues(intRI))
                                Case "11"
                                    intMixSwitch(1) = intMixSwitch(1) + Int(strUseValues(intRI))
                                Case "66"   'RM1809032_66追加
                                    intMixSwitch(8) = intMixSwitch(8) + Int(strUseValues(intRI))
                            End Select
                        ElseIf Left(strKataValues(intRI), 4) = "NW4G" Then
                            Select Case strKataValues(intRI).Substring(6, 2)
                                Case "10"
                                    intMixSwitch(2) = intMixSwitch(2) + Int(strUseValues(intRI))
                                Case "20"
                                    intMixSwitch(3) = intMixSwitch(3) + Int(strUseValues(intRI))
                                Case "30"
                                    intMixSwitch(4) = intMixSwitch(4) + Int(strUseValues(intRI))
                                Case "40"
                                    intMixSwitch(5) = intMixSwitch(5) + Int(strUseValues(intRI))
                                Case "50"
                                    intMixSwitch(6) = intMixSwitch(6) + Int(strUseValues(intRI))
                            End Select
                        End If
                    End If
                    If Len(Trim(strKataValues(intRI))) > 8 Then
                        If strKataValues(intRI).Substring(7, 2) = "MP" Then
                            intMixSwitch(7) = intMixSwitch(7) + Int(strUseValues(intRI))
                        End If
                    End If
                End If

                intElectCnt = intElectCnt + Int(strUseValues(intRI))
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    '********************************************************************************************
    '*【関数名】
    '*   fncFoundValve
    '*【処理】
    '*   バルブブロック選択列存在チェック
    '*【引数】
    '   intProc     処理区分[0: 7.11.1  7.11.2 (右→左)][1: 7.11.1(左→右)][2: 7.11.2 (左→右)]
    '   intStNo     処理開始列No                    intEdNo     処理終了列No
    '   arySelInfo  選択セル情報配列
    '********************************************************************************************
    Public Shared Function fncFoundValve(objKtbnStrc As KHKtbnStrc, ByVal intProc As Integer, _
                                         ByVal intStNo As Integer, ByVal intEdNo As Integer) As Boolean
        Dim intLoop As Integer
        Dim intStep As Integer
        Dim bolReturn As Boolean = False
        Try
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo
            If intProc = 0 Then
                intStep = -1
            Else
                intStep = 1
            End If
            intLoop = intStNo
            Do Until intLoop = intEdNo - 1
                If intProc = 1 Then
                    For intRI As Integer = Siyou_15.Partition1 - 1 To Siyou_15.Partition2 - 1
                        If arySelectInf(intRI)(intLoop) = "1" Then
                            Exit Do
                        End If
                    Next
                End If
                For intRI As Integer = Siyou_15.Valve1 - 1 To Siyou_15.Valve8 - 1
                    If arySelectInf(intRI)(intLoop) = "1" Then
                        bolReturn = True
                        Exit Do
                    End If
                Next
                intLoop = intLoop + intStep
            Loop
            fncFoundValve = bolReturn
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncFoundExhaust
    '*【処理】
    '*   給排気ブロック選択列存在チェック
    '*【引数】
    '   intProc     処理区分[0: 7.11.4 7.11.3 (右→左)][1: 7.11.4(左→右)]
    '   intStNo     処理開始列No                    intEdNo         処理終了列No
    '   arySelInfo  選択セル情報配列                strSelKataban   選択形番配列
    '********************************************************************************************
    Public Shared Function fncFoundExhaust(objKtbnStrc As KHKtbnStrc, ByVal intProc As Integer, _
                                           ByVal intStNo As Integer, ByVal intEdNo As Integer) As Boolean
        Dim intLoop As Integer
        Dim intStep As Integer
        Dim bolReturn As Boolean = False
        Try
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            If intProc = 1 Then
                intStep = 1
            Else
                intStep = -1
            End If
            intLoop = intStNo
            Do Until intLoop = intEdNo - 1
                If intProc <> 0 Then
                    For intRI As Integer = Siyou_15.Partition1 - 1 To Siyou_15.Partition2 - 1
                        If arySelectInf(intRI)(intLoop) = "1" Then
                            Exit Do
                        End If
                    Next
                End If
                For intRI As Integer = Siyou_15.Exhaust1 - 1 To Siyou_15.Exhaust2 - 1
                    If arySelectInf(intRI)(intLoop) = "1" Then
                        bolReturn = True
                        Exit Do
                    End If
                Next
                intLoop = intLoop + intStep
            Loop
            fncFoundExhaust = bolReturn
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncCheckExhaust
    '*【処理】
    '*   給排気ブロック選択形番チェック
    '*【引数】
    '   intProc     処理区分[0: 7.13(右→左)][1: 7.13(左→右)][2: 7.12(右→左)]
    '   intStNo     処理開始列No                    intEdNo         処理終了列No
    '   arySelInfo  選択セル情報配列                strSelKataban   選択形番配列
    '********************************************************************************************
    Public Shared Function fncCheckExhaust(objKtbnStrc As KHKtbnStrc, ByVal intProc As Integer, _
                                           ByVal intStNo As Integer, ByVal intEdNo As Integer) As Boolean
        Dim intLoop As Integer
        Dim intStep As Integer
        Dim bolReturn As Boolean = False
        Try
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo
            If intProc = 1 Then
                intStep = 1
            Else
                intStep = -1
            End If
            intLoop = intStNo
            Do Until intLoop = intEdNo - 1
                '区分 = 2 (仕切りブロック/給排気ブロックチェック)の場合、先に仕切りブロックの選択を確認する
                If intProc = 2 Then
                    For intRI As Integer = Siyou_15.Partition1 - 1 To Siyou_15.Partition2 - 1
                        If arySelectInf(intRI)(intLoop) = "1" Then
                            Exit Do
                        End If
                    Next
                End If

                For intRI As Integer = Siyou_15.Exhaust1 - 1 To Siyou_15.Exhaust2 - 1
                    If arySelectInf(intRI)(intLoop) = "1" Then
                        If strKataValues(intRI).Contains("X") Then
                            If intProc = 2 Or Not strKataValues(intRI).Contains("-S") Then
                                bolReturn = True
                            End If
                        End If
                        If intProc = 2 Then
                            If bolReturn Then
                                Exit Do
                            End If
                        Else
                            Exit Do
                        End If
                    End If
                Next
                For intRI As Integer = Siyou_15.Partition1 - 1 To Siyou_15.Partition2 - 1
                    If arySelectInf(intRI)(intLoop) = "1" Then
                        Exit Do
                    End If
                Next
                intLoop = intLoop + intStep
            Loop
            fncCheckExhaust = bolReturn
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function
End Class
