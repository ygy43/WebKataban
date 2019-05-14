Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_18

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
        Dim intSolCnt As Integer = 0
        Dim intColR As Integer = 0
        Dim bolChkFlag As Boolean = False
        Dim intMixSwtchCnt As Integer = 0
        Dim intMixConCnt As Integer = 0
        Dim intElectSeq As Integer
        Dim intColCnt As Integer = 40       'RM1803032_マニホールド連数拡張
        Dim intPosRowCnt As Integer = 21    'RM1803032_スペーサ行追加
        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban
Dim intNo As Integer = 0

        fncInputChk = False

        Select Case objKtbnStrc.strcSelection.strKeyKataban
            Case "R", "U", "S", "V"
                intNo = 10
            Case Else
                intNo = 9
        End Select

        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            '----- コントロール値変更時チェック --------------------------------------------------
            '形番選択チェック
            For inti As Integer = 0 To intPosRowCnt - 1
                If strKataValues(inti) = String.Empty And strUseValues(inti) > 0 Then
                    sbCoordinates.Append(CStr(inti + 1) & strComma & "0")
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1400"
                    Exit Function
                End If
            Next

            'ブランクプラグ＆サイレンサ・検査成績書＆ケーブル数値テキスト入力値チェック
            If Not SiyouBLL.fncOtherKataCheck(objKtbnStrc, Siyou_18.Plug1, Siyou_18.Inspect2, _
                                     Siyou_18.Rail, strMsgCd) Then
                Exit Function
            End If

            '重複チェック
            If Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_18.Elect1, Siyou_18.Elect8) Or _
               Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_18.Spacer1, Siyou_18.Spacer4) Or _
               Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_18.Exhaust1, Siyou_18.Exhaust3) Or _
               Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_18.Partition1, Siyou_18.Partition2) Or _
               Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_18.EndLeft, Siyou_18.EndRight) Or _
               Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_18.Plug1, Siyou_18.Plug3) Or _
               Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_18.Inspect1, Siyou_18.Inspect2) Then
                strMsgCd = "W1330"
                Exit Function
            End If

            '取付レール長さ入力値チェック
            If strKataValues(Siyou_18.Rail - 1).ToString.Length <= 0 Then strKataValues(Siyou_18.Rail - 1) = 0
            If Not SiyouBLL.fncRailchk(strKataValues(Siyou_18.Rail - 1), strUseValues(Siyou_18.Rail - 1), dblStdNum, strMsgCd) Then
                strMsg = Siyou_18.Rail & ",0"
                Exit Function
            End If

            '----- 入力内容値チェック --------------------------------------------------
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

            '縦列複数チェック
            Dim bolSelFlag1 As Boolean
            Dim bolSelFlag2 As Boolean
            bolSelFlag1 = False
            bolSelFlag2 = False
            For intCI As Integer = 1 To intColR
                For intRI As Integer = Siyou_18.Equip To Siyou_18.EndRight
                    If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                        If intRI = Siyou_18.Spacer1 Or intRI = Siyou_18.Spacer2 Or intRI = Siyou_18.Spacer3 Or intRI = Siyou_18.Spacer4 Then
                            'スペーサ(１１～１２行目)が同列で複数選択されていた場合、エラー
                            If bolSelFlag2 = True Then
                                sbCoordinates.Append("0" & strComma & CStr(intCI))
                                strMsg = sbCoordinates.ToString
                                strMsgCd = "W1390"
                                Exit Function
                            Else
                                bolSelFlag2 = True
                            End If
                        Else
                            'スペーサ(１１～１２行目)以外が同列で複数選択されていた場合、エラー
                            If bolSelFlag1 = True Then
                                sbCoordinates.Append("0" & strComma & CStr(intCI))
                                strMsg = sbCoordinates.ToString
                                strMsgCd = "W1390"
                                Exit Function
                            Else
                                bolSelFlag1 = True
                            End If
                        End If
                    End If
                Next
                bolSelFlag1 = False
                bolSelFlag2 = False
            Next


            '7.1 接続位置チェック
            '一つも未選択の場合、エラー
            If intColR = 0 Then
                strMsgCd = "W1030"
                Exit Function
            End If
            '最右列まで連続チェックされていない場合、エラー
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

            '7.2 電装ブロック複数指定チェック
            If strUseValues(Siyou_18.Equip - 1) > 1 Then
                sbCoordinates.Append(CStr(Siyou_18.Equip) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W1640"
                Exit Function
            End If

            '7.3 エンドブロック(左)複数指定チェック
            If strUseValues(Siyou_18.EndLeft - 1) > 1 Then
                sbCoordinates.Append(CStr(Siyou_18.EndLeft) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W1100"
                Exit Function
            End If

            '7.4 エンドブロック(右)複数指定チェック
            If strUseValues(Siyou_18.EndRight - 1) > 1 Then
                sbCoordinates.Append(CStr(Siyou_18.EndRight) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W1100"
                Exit Function
            End If

            Dim strDen As String = String.Empty
            Dim strLen As String = String.Empty
            Dim strKoukei As String = String.Empty
            Dim strOption As String = String.Empty
            If objKtbnStrc.strcSelection.strOpSymbol(3).ToString.Trim = "R" Then
                strDen = objKtbnStrc.strcSelection.strOpSymbol(5)
                strLen = objKtbnStrc.strcSelection.strOpSymbol(9)
                strOption = objKtbnStrc.strcSelection.strOpSymbol(8)
                strKoukei = objKtbnStrc.strcSelection.strOpSymbol(4)
            Else
                strKoukei = objKtbnStrc.strcSelection.strOpSymbol(3)
                strDen = objKtbnStrc.strcSelection.strOpSymbol(4)
                strLen = objKtbnStrc.strcSelection.strOpSymbol(8)
                strOption = objKtbnStrc.strcSelection.strOpSymbol(7)
            End If

            '7.5 エンドブロック形番選択チェック
            '電線／省配線接続左１文字が"T"
            If Left(strDen.ToString, 1) = "T" Then
                '電線／省配線接続４文字目が"R"
                If Left(strDen.ToString & Space(4), 4).Substring(3, 1) = "R" Then
                    '１７行目形番未選択エラー
                    If Trim(strKataValues(Siyou_18.EndLeft - 1)).Length = 0 Then
                        sbCoordinates.Append(CStr(Siyou_18.EndLeft) & strComma & "0")
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W1620"
                        Exit Function
                    End If
                Else    '電線／省配線接続４文字目が"R"以外
                    '１８行目形番未選択エラー
                    If Trim(strKataValues(Siyou_18.EndRight - 1)).Length = 0 Then
                        sbCoordinates.Append(CStr(Siyou_18.EndRight) & strComma & "0")
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W2070"
                        Exit Function
                    End If
                End If
            Else        '電線／省配線接続左１文字が"T"以外
                '１７行目形番未選択エラー
                If Trim(strKataValues(Siyou_18.EndLeft - 1)).Length = 0 Then
                    sbCoordinates.Append(CStr(Siyou_18.EndLeft) & strComma & "0")
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1620"
                    Exit Function
                End If
                '１８行目形番未選択エラー
                If Trim(strKataValues(Siyou_18.EndRight - 1)).Length = 0 Then
                    sbCoordinates.Append(CStr(Siyou_18.EndRight) & strComma & "0")
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W2070"
                    Exit Function
                End If
            End If

            '7.6 電装ブロック・エンドブロック必須チェック
            '電線／省配線接続の左１文字が"T"
            If Left(strDen.ToString, 1) = "T" Then
                '電線／省配線接続の４文字目が"R"
                If Left(strDen.ToString & Space(4), 4).Substring(3, 1) = "R" Then
                    '最右列が１行目以外エラー
                    If arySelectInf(Siyou_18.Equip - 1)(intColR - 1) = "1" Then
                    Else
                        sbCoordinates.Append("0" & strComma & CStr(intColR + 1))
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W2080"
                        Exit Function
                    End If
                Else        '電線／省配線接続の４文字目が"R"以外
                    '最右列が１８行目以外エラー
                    If arySelectInf(Siyou_18.EndRight - 1)(intColR - 1) = "1" Then
                    Else
                        sbCoordinates.Append(Siyou_18.EndRight & strComma & CStr(intColR + 1))
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W1650"
                        Exit Function
                    End If
                End If
            Else        '電線／省配線接続の左１文字が"T"
                '最右列が１８行目以外エラー
                If arySelectInf(Siyou_18.EndRight - 1)(intColR - 1) = "1" Then
                Else
                    sbCoordinates.Append(Siyou_18.EndRight & strComma & CStr(intColR + 1))
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1650"
                    Exit Function
                End If
            End If

            '7.7 ソレノイド点数チェック
            If Left(strDen.ToString, 1) = "T" Then
                intSolCnt = 0
                'ソレノイドカウント取得(２～９行目)
                For intRI As Integer = Siyou_18.Elect1 - 1 To Siyou_18.Elect8 - 1
                    If strUseValues(intRI) > 0 Then
                        If InStr(strKataValues(intRI), "-MP") = 0 Then
                            If Left(strKataValues(intRI) & Space(6), 6).Substring(5, 1) = "1" Then
                                intSolCnt = intSolCnt + CInt(strUseValues(intRI))
                            ElseIf Left(strKataValues(intRI) & Space(6), 6).Substring(5, 1) = "-" Then
                            Else
                                intSolCnt = intSolCnt + CInt(strUseValues(intRI)) * 2
                            End If
                        Else
                            If Left(strKataValues(intRI) & Space(intNo), intNo).Substring(intNo - 1, 1) = "S" Then
                                intSolCnt = intSolCnt + CInt(strUseValues(intRI))
                            ElseIf Left(strKataValues(intRI) & Space(intNo), intNo).Substring(intNo - 1, 1) = "D" Then
                                intSolCnt = intSolCnt + CInt(strUseValues(intRI)) * 2
                            End If
                        End If
                    End If
                Next
                'ソレノイドカウント＞ソレノイドMAX時、エラー
                If intSolCnt > KHKataban.fncGetMaxSol(objKtbnStrc.strcSelection.strOpSymbol, 18) Then
                    strMsgCd = "W1150"
                    Exit Function
                End If
            End If

            '7.8 バルブブロック使用数チェック
            intElectSeq = 0

            'バルブブロック(２～９行目)について確認
            For intRI As Integer = Siyou_18.Elect1 - 1 To Siyou_18.Elect8 - 1
                '形番要素が選択かつ使用数 > 0の場合
                If strKataValues(intRI).Length > 0 And CInt(strUseValues(intRI)) > 0 Then
                    '電磁弁連数カウント
                    intElectSeq = intElectSeq + CInt(strUseValues(intRI))
                End If
            Next

            '電磁弁エラーチェック
            '電磁弁連数 > 最大連数の場合、エラー
            If intElectSeq > CInt(strLen.ToString) Then
                strMsgCd = "W1170"
                Exit Function
            End If
            '電磁弁連数 < 最大連数の場合、エラー
            If intElectSeq < CInt(strLen.ToString) Then
                strMsgCd = "W1180"
                Exit Function
            End If
            '位置切替区分の頭が"8"の場合のみ
            If Left(objKtbnStrc.strcSelection.strOpSymbol(1).ToString, 1) = "8" And InStr(1, strSeriesKata, "X12") = 0 Then
                Dim flgH As Boolean = False
                If strOption.ToString = "H" Then flgH = True
                If Not SiyouBLL.fncMixSwtchCheck(objKtbnStrc, Siyou_18.Elect1 - 1, Siyou_18.Elect8 - 1, flgH, strMsgCd) Then
                    Exit Function
                End If
            End If
            '接続口径左2文字が"CX"の場合
            If Left(strKoukei.ToString, 2) = "CX" Then
                'ミックスチェック値(接続口径)のTrueの値が１つ以下の場合、エラー
                If Not SiyouBLL.fncMixBlockCheck(objKtbnStrc, Siyou_18.Elect1 - 1, Siyou_18.Elect8 - 1, strMsgCd) Then
                    Exit Function
                End If
            End If

            '7.10 スペーサ・バルブブロック組合せチェック
            Dim bolCmbCheck As Boolean = False
            Dim bolSelChk As Boolean = False
            For intCI As Integer = 1 To intColR
                For intRI1 As Integer = Siyou_18.Spacer1 To Siyou_18.Spacer4
                    If arySelectInf(intRI1 - 1)(intCI - 1) = "1" Then
                        For intRI2 As Integer = Siyou_18.Elect1 To Siyou_18.Elect8
                            If arySelectInf(intRI2 - 1)(intCI - 1) = "1" Then
                                bolCmbCheck = True
                                Exit For
                            End If
                        Next
                        If bolCmbCheck Then
                            Exit For
                        End If
                        bolSelChk = True
                    End If
                Next
            Next
            'スペーサと同列にバルブブロックが選択されていない場合、エラー
            If bolCmbCheck = False And bolSelChk Then
                sbCoordinates.Append("0" & strComma & Siyou_18.Spacer1 & "|0," & Siyou_18.Spacer2 & "|0," & Siyou_18.Spacer3 & "|0," & Siyou_18.Spacer4)
                strMsg = sbCoordinates.ToString
                strMsgCd = "W2250"
                Exit Function
            End If
            bolCmbCheck = False

            '7.11 仕切ブロック・給排気ブロック・エンドブロック組合せチェック
            '給排気に'X'が含まれる時、エンドブロックに"X"が含まれない場合、エラー
            If (CInt(strUseValues(Siyou_18.Exhaust1 - 1)) > 0 And InStr(strKataValues(Siyou_18.Exhaust1 - 1), "X") = 0) Or _
               (CInt(strUseValues(Siyou_18.Exhaust2 - 1)) > 0 And InStr(strKataValues(Siyou_18.Exhaust2 - 1), "X") = 0) Or _
               (CInt(strUseValues(Siyou_18.Exhaust3 - 1)) > 0 And InStr(strKataValues(Siyou_18.Exhaust3 - 1), "X") = 0) Or _
               (CInt(strUseValues(Siyou_18.Exhaust1 - 1)) = 0 And CInt(strUseValues(Siyou_18.Exhaust2 - 1)) = 0 And CInt(strUseValues(Siyou_18.Exhaust3 - 1)) = 0) Then
            Else
                If (CInt(strUseValues(Siyou_18.EndLeft - 1)) > 0 And InStr(strKataValues(Siyou_18.EndLeft - 1), "X") > 0) Or _
                   (CInt(strUseValues(Siyou_18.EndRight - 1)) > 0 And InStr(strKataValues(Siyou_18.EndRight - 1), "X") > 0) Then
                Else
                    strMsgCd = "W2260"
                    Exit Function
                End If
            End If

            Dim bolPartPos As Boolean               '仕切ブロック設置位置チェック
            Dim intPartPos As Integer               '仕切ブロック設置位置
            Dim bolExhaustChk As Boolean = False    '給排気ブロック選択チェック
            '仕切ブロック未選択
            If CInt(strUseValues(Siyou_18.Partition1 - 1)) = 0 And CInt(strUseValues(Siyou_18.Partition2 - 1)) = 0 Then
                '給排気未選択時
                If CInt(strUseValues(Siyou_18.Exhaust1 - 1)) = 0 And CInt(strUseValues(Siyou_18.Exhaust2 - 1)) = 0 And CInt(strUseValues(Siyou_18.Exhaust3 - 1)) = 0 Then
                    sbCoordinates.Append(CStr(Siyou_18.Exhaust1) & strComma & "0")
                    sbCoordinates.Append(strPipe & CStr(Siyou_18.Exhaust2) & strComma & "0")
                    sbCoordinates.Append(strPipe & CStr(Siyou_18.Exhaust3) & strComma & "0")
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1700"
                    Exit Function
                End If
            Else
                '仕切ブロック選択あり
                For intCI1 As Integer = 1 To intColR
                    '給排気選択時
                    '仕切ブロック設置位置取得
                    intPartPos = 0
                    If arySelectInf(Siyou_18.Partition1 - 1)(intCI1 - 1) = "1" Or _
                       arySelectInf(Siyou_18.Partition2 - 1)(intCI1 - 1) = "1" Then
                        intPartPos = intCI1
                    End If

                    If intPartPos > 0 Then
                        '仕切ブロック左チェック
                        bolPartPos = False
                        For intCI2 As Integer = intPartPos - 1 To 1 Step -1
                            '仕切ブロックが選択されている場合、確認処理終了
                            For intRI As Integer = Siyou_18.Partition1 To Siyou_18.Partition2
                                If arySelectInf(intRI - 1)(intCI2 - 1) = "1" Then
                                    Exit For
                                End If
                            Next
                            '仕切ブロック左列にバルブブロックが選択されている場合、確認終了
                            If arySelectInf(Siyou_18.Elect1 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Elect2 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Elect3 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Elect4 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Elect5 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Elect6 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Elect7 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Elect8 - 1)(intCI2 - 1) = "1" Then
                                bolPartPos = True
                                Exit For
                            End If
                        Next
                        'バルブブロック選択チェックがFalseの場合、エラー
                        If bolPartPos = False Then
                            If arySelectInf(Siyou_18.Partition1 - 1)(intCI1 - 1) = "1" Then
                                sbCoordinates.Append(CStr(Siyou_18.Partition1) & strComma & CStr(intCI1))
                            ElseIf arySelectInf(Siyou_18.Partition2 - 1)(intCI1 - 1) = "1" Then
                                sbCoordinates.Append(CStr(Siyou_18.Partition2) & strComma & CStr(intCI1))
                            End If
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W1710"
                            Exit Function
                        End If

                        '仕切ブロック右チェック
                        bolPartPos = False
                        For intCI2 As Integer = intPartPos + 1 To intColR
                            '仕切ブロックが選択されている場合、確認処理終了
                            For intRI As Integer = Siyou_18.Partition1 To Siyou_18.Partition2
                                If arySelectInf(intRI - 1)(intCI2 - 1) = "1" Then
                                    Exit For
                                End If
                            Next
                            '仕切ブロック右列にバルブブロックが選択されている場合、確認終了
                            If arySelectInf(Siyou_18.Elect1 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Elect2 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Elect3 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Elect4 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Elect5 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Elect6 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Elect7 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Elect8 - 1)(intCI2 - 1) = "1" Then
                                bolPartPos = True
                                Exit For
                            End If
                        Next
                        'バルブブロック選択チェックがFalseの場合、エラー
                        If bolPartPos = False Then
                            If arySelectInf(Siyou_18.Partition1 - 1)(intCI1 - 1) = "1" Then
                                sbCoordinates.Append(CStr(Siyou_18.Partition1) & strComma & CStr(intCI1))
                            ElseIf arySelectInf(Siyou_18.Partition2 - 1)(intCI1 - 1) = "1" Then
                                sbCoordinates.Append(CStr(Siyou_18.Partition2) & strComma & CStr(intCI1))
                            End If
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W1710"
                            Exit Function
                        End If

                        '給排気ブロック左チェック
                        bolExhaustChk = False
                        For intCI2 As Integer = intPartPos - 1 To 1 Step -1
                            '仕切ブロックが選択されている場合、確認処理終了
                            If arySelectInf(Siyou_18.Partition1 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Partition2 - 1)(intCI2 - 1) = "1" Then
                                Exit For
                            End If
                            'スペーサ～給排気ブロックが選択されている場合、給排気ブロック選択チェックをTrueにして確認処理終了
                            For intRI As Integer = Siyou_18.Spacer1 To Siyou_18.Exhaust3
                                If arySelectInf(intRI - 1)(intCI2 - 1) = "1" Then
                                    bolExhaustChk = True
                                    Exit For
                                End If
                            Next
                            If bolExhaustChk = True Then
                                Exit For
                            End If
                        Next
                        '給排気ブロック選択チェックがFalseの場合、エラー
                        If bolExhaustChk = False Then
                            'エンドブロックに'X'を含む場合はエラー対象外
                            If CInt(strUseValues(Siyou_18.EndLeft - 1)) > 0 And InStr(strKataValues(Siyou_18.EndLeft - 1), "X") > 0 Then
                            Else
                                If arySelectInf(Siyou_18.Partition1 - 1)(intCI1 - 1) = "1" Then
                                    sbCoordinates.Append(CStr(Siyou_18.Partition1) & strComma & CStr(intCI1))
                                ElseIf arySelectInf(Siyou_18.Partition2 - 1)(intCI1 - 1) = "1" Then
                                    sbCoordinates.Append(CStr(Siyou_18.Partition2) & strComma & CStr(intCI1))
                                End If
                                strMsg = sbCoordinates.ToString
                                strMsgCd = "W2270"
                                Exit Function
                            End If
                        End If

                        '給排気ブロック右チェック
                        bolExhaustChk = False
                        For intCI2 As Integer = intPartPos + 1 To intColR
                            '仕切ブロックが選択されている場合、確認処理終了
                            If arySelectInf(Siyou_18.Partition1 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Partition2 - 1)(intCI2 - 1) = "1" Then
                                Exit For
                            End If
                            'スペーサ～給排気ブロックが選択されている場合、給排気ブロック選択チェックをTrueにして確認処理終了
                            For intRI As Integer = Siyou_18.Spacer1 To Siyou_18.Exhaust3
                                If arySelectInf(intRI - 1)(intCI2 - 1) = "1" Then
                                    bolExhaustChk = True
                                    Exit For
                                End If
                            Next
                            If bolExhaustChk = True Then
                                Exit For
                            End If
                        Next
                        '給排気ブロック選択チェックがFalseの場合、エラー
                        If bolExhaustChk = False Then
                            'エンドブロックに'X'を含む場合はエラー対象外
                            If CInt(strUseValues(Siyou_18.EndRight - 1)) > 0 And InStr(strKataValues(Siyou_18.EndRight - 1), "X") > 0 Then
                            Else
                                If arySelectInf(Siyou_18.Partition1 - 1)(intCI1 - 1) = "1" Then
                                    sbCoordinates.Append(CStr(Siyou_18.Partition1) & strComma & CStr(intCI1))
                                ElseIf arySelectInf(Siyou_18.Partition2 - 1)(intCI1 - 1) = "1" Then
                                    sbCoordinates.Append(CStr(Siyou_18.Partition2) & strComma & CStr(intCI1))
                                End If
                                strMsg = sbCoordinates.ToString
                                strMsgCd = "W2270"
                                Exit Function
                            End If
                        End If
                    End If
                Next
            End If

            '仕切ブロック・給排気ブロック・エンドブロック組合せチェック２
            Dim bolAutoSelChk As Boolean = False        '大気開放選択チェック
            Dim bolPartLChk As Boolean = False          '仕切ブロック左チェック
            Dim bolPartRChk As Boolean = False          '仕切ブロック右チェック
            Dim strPosKataban As String
            For intCI1 As Integer = 1 To intColR
                intPartPos = 0
                strPosKataban = ""
                '仕切ブロック設置位置取得
                For intRI As Integer = Siyou_18.Partition1 To Siyou_18.Partition2
                    If arySelectInf(intRI - 1)(intCI1 - 1) = "1" Then
                        intPartPos = intCI1
                        strPosKataban = strKataValues(intRI - 1)
                        Exit For
                    End If
                Next

                If intPartPos > 0 Then
                    '仕切ブロック形番要素が"-SP"以外
                    If InStr(strPosKataban, "-SP") = 0 Then
                        '仕切ブロック左チェック
                        bolAutoSelChk = False
                        For intCI2 As Integer = intPartPos - 1 To 1 Step -1
                            '仕切ブロックが選択されている場合、確認処理終了
                            If arySelectInf(Siyou_18.Partition1 - 1)(intCI2 - 1) = "1" And InStr(strKataValues(Siyou_18.Partition1 - 1), "-SP") = 0 Or _
                               arySelectInf(Siyou_18.Partition2 - 1)(intCI2 - 1) = "1" And InStr(strKataValues(Siyou_18.Partition2 - 1), "-SP") = 0 Then
                                Exit For
                            End If
                            '給排気ブロックが選択されている場合、大気開放チェックをTrueにして確認処理終了
                            For intRI As Integer = Siyou_18.Exhaust1 To Siyou_18.Exhaust3
                                If arySelectInf(intRI - 1)(intCI2 - 1) = "1" And _
                                   InStr(strKataValues(intRI - 1), "X") = 0 Then
                                    bolAutoSelChk = True
                                    Exit For
                                End If
                            Next
                            If bolAutoSelChk = True Then
                                Exit For
                            End If
                        Next
                        '大気開放選択チェックがFalseの場合、エラー
                        If bolAutoSelChk = False Then
                            'エンドブロックに'X'を含む場合はエラー対象外
                            If CInt(strUseValues(Siyou_18.EndLeft - 1)) > 0 And InStr(strKataValues(Siyou_18.EndLeft - 1), "X") > 0 Then
                            Else
                                If arySelectInf(Siyou_18.Partition1 - 1)(intCI1 - 1) = "1" Then
                                    sbCoordinates.Append(CStr(Siyou_18.Partition1) & strComma & CStr(intCI1))
                                ElseIf arySelectInf(Siyou_18.Partition2 - 1)(intCI1 - 1) = "1" Then
                                    sbCoordinates.Append(CStr(Siyou_18.Partition2) & strComma & CStr(intCI1))
                                End If
                                strMsg = sbCoordinates.ToString
                                strMsgCd = "W2260"
                                Exit Function
                            End If
                        End If

                        '仕切ブロック右チェック
                        bolAutoSelChk = False
                        '仕切ブロックが選択されている場合、確認処理終了
                        For intCI2 As Integer = intPartPos + 1 To intColR
                            If arySelectInf(Siyou_18.Partition1 - 1)(intCI2 - 1) = "1" And InStr(strKataValues(Siyou_18.Partition1 - 1), "-SP") = 0 Or _
                               arySelectInf(Siyou_18.Partition2 - 1)(intCI2 - 1) = "1" And InStr(strKataValues(Siyou_18.Partition2 - 1), "-SP") = 0 Then
                                Exit For
                            End If
                            '給排気ブロックが選択されている場合、大気開放チェックをTrueにして確認処理終了
                            For intRI As Integer = Siyou_18.Exhaust1 To Siyou_18.Exhaust3
                                If arySelectInf(intRI - 1)(intCI2 - 1) = "1" And _
                                   InStr(strKataValues(intRI - 1), "X") = 0 Then
                                    bolAutoSelChk = True
                                    Exit For
                                End If
                            Next
                            If bolAutoSelChk = True Then
                                Exit For
                            End If
                        Next

                        '大気開放選択チェックがFalseの場合、エラー
                        If bolAutoSelChk = False Then
                            'エンドブロックに'X'を含む場合はエラー対象外
                            If CInt(strUseValues(Siyou_18.EndRight - 1)) > 0 And InStr(strKataValues(Siyou_18.EndRight - 1), "X") > 0 Then
                            Else
                                If arySelectInf(Siyou_18.Partition1 - 1)(intCI1 - 1) = "1" Then
                                    sbCoordinates.Append(CStr(Siyou_18.Partition1) & strComma & CStr(intCI1))
                                ElseIf arySelectInf(Siyou_18.Partition2 - 1)(intCI1 - 1) = "1" Then
                                    sbCoordinates.Append(CStr(Siyou_18.Partition2) & strComma & CStr(intCI1))
                                End If
                                strMsg = sbCoordinates.ToString
                                strMsgCd = "W2260"
                                Exit Function
                            End If
                        End If

                        '仕切ブロック左／右チェック
                        bolPartLChk = False
                        bolPartRChk = False
                        '仕切ブロックより左側を確認
                        For intCI2 As Integer = intPartPos - 1 To 1 Step -1
                            '仕切ブロックが選択されている場合、確認処理終了
                            If arySelectInf(Siyou_18.Partition1 - 1)(intCI2 - 1) = "1" And InStr(strKataValues(Siyou_18.Partition1 - 1), "-SP") = 0 Or _
                               arySelectInf(Siyou_18.Partition2 - 1)(intCI2 - 1) = "1" And InStr(strKataValues(Siyou_18.Partition2 - 1), "-SP") = 0 Then
                                Exit For
                            End If
                            '給排気ブロックが選択されている場合、仕切ブロック左チェックをTrueにして確認処理終了
                            For intRI As Integer = Siyou_18.Exhaust1 To Siyou_18.Exhaust3
                                If arySelectInf(intRI - 1)(intCI2 - 1) = "1" And _
                                   InStr(strKataValues(intRI - 1), "X") = 0 Then
                                    bolPartLChk = True
                                    Exit For
                                End If
                            Next
                            If bolPartLChk = True Then
                                Exit For
                            End If
                        Next
                        '仕切ブロックより右側を確認
                        For intCI2 As Integer = intPartPos + 1 To intColR
                            '仕切ブロックが選択されている場合、確認処理終了
                            If arySelectInf(Siyou_18.Partition1 - 1)(intCI2 - 1) = "1" And InStr(strKataValues(Siyou_18.Partition1 - 1), "-SP") = 0 Or _
                               arySelectInf(Siyou_18.Partition2 - 1)(intCI2 - 1) = "1" And InStr(strKataValues(Siyou_18.Partition2 - 1), "-SP") = 0 Then
                                Exit For
                            End If
                            '給排気ブロックが選択されている場合、仕切ブロック右チェックをTrueにして確認処理終了
                            For intRI As Integer = Siyou_18.Exhaust1 To Siyou_18.Exhaust3
                                If arySelectInf(intRI - 1)(intCI2 - 1) = "1" And _
                                   InStr(strKataValues(intRI - 1), "X") = 0 Then
                                    bolPartRChk = True
                                    Exit For
                                End If
                            Next
                            If bolPartRChk = True Then
                                Exit For
                            End If
                        Next
                        '仕切ブロック左／右チェック共にFalseの場合、エラー
                        If bolPartLChk = False And bolPartRChk = False Then
                            'エンドブロックに'X'が含まれる場合、エラー対象外
                            If CInt(strUseValues(Siyou_18.EndLeft - 1)) > 0 And InStr(strKataValues(Siyou_18.EndLeft - 1), "X") > 0 Or _
                               CInt(strUseValues(Siyou_18.EndRight - 1)) > 0 And InStr(strKataValues(Siyou_18.EndRight - 1), "X") > 0 Then
                            Else
                                If arySelectInf(Siyou_18.Partition1 - 1)(intCI1 - 1) = "1" Then
                                    sbCoordinates.Append(CStr(Siyou_18.Partition1) & strComma & CStr(intCI1))
                                ElseIf arySelectInf(Siyou_18.Partition2 - 1)(intCI1 - 1) = "1" Then
                                    sbCoordinates.Append(CStr(Siyou_18.Partition2) & strComma & CStr(intCI1))
                                End If
                                strMsg = sbCoordinates.ToString
                                strMsgCd = "W2260"
                                Exit Function
                            End If
                        End If
                    Else
                        '対象列形番要素に"-SP"を含む場合
                        '仕切ブロック左チェック
                        bolPartLChk = False
                        For intCI2 As Integer = intPartPos - 1 To 1 Step -1
                            If arySelectInf(Siyou_18.Partition1 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Partition2 - 1)(intCI2 - 1) = "1" Then
                                Exit For
                            End If
                            For intri As Integer = Siyou_18.Exhaust1 To Siyou_18.Exhaust3
                                If arySelectInf(intri - 1)(intCI2 - 1) = "1" And _
                                   InStr(strKataValues(intri - 1), "X") = 0 Then
                                    bolPartLChk = True
                                    Exit For
                                End If
                            Next
                        Next
                        '仕切ブロック左チェック
                        bolPartRChk = False
                        For intCI2 As Integer = intPartPos + 1 To intColR
                            If arySelectInf(Siyou_18.Partition1 - 1)(intCI2 - 1) = "1" Or _
                               arySelectInf(Siyou_18.Partition2 - 1)(intCI2 - 1) = "1" Then
                                Exit For
                            End If
                            For intri As Integer = Siyou_18.Exhaust1 To Siyou_18.Exhaust3
                                If arySelectInf(intri - 1)(intCI2 - 1) = "1" And _
                                   InStr(strKataValues(intri - 1), "X") = 0 Then
                                    bolPartRChk = True
                                    Exit For
                                End If
                            Next
                        Next

                        '仕切ブロック左／右チェック
                        If bolPartLChk = False And bolPartRChk = False Then
                            If CInt(strUseValues(Siyou_18.EndLeft - 1)) > 0 And InStr(strKataValues(Siyou_18.EndLeft - 1), "X") > 0 Or _
                               CInt(strUseValues(Siyou_18.EndRight - 1)) > 0 And InStr(strKataValues(Siyou_18.EndRight - 1), "X") > 0 Then
                            Else
                                If arySelectInf(Siyou_18.Partition1 - 1)(intCI1 - 1) = "1" Then
                                    sbCoordinates.Append(CStr(Siyou_18.Partition1) & strComma & CStr(intCI1))
                                ElseIf arySelectInf(Siyou_18.Partition2 - 1)(intCI1 - 1) = "1" Then
                                    sbCoordinates.Append(CStr(Siyou_18.Partition2) & strComma & CStr(intCI1))
                                End If
                                strMsg = sbCoordinates.ToString
                                strMsgCd = "W2260"
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next

            'バルブブロックチェック
            If Not SiyouBLL.fncBlockCheck(strUseValues, strKataValues, Siyou_18.Elect1 - 1, Siyou_18.Elect8 - 1, strMsgCd) Then
                Exit Function
            End If

            '給気スペーサオプションチェック
            If strOption.ToString = "Z1" Then
                If CInt(strUseValues(Siyou_18.Spacer1 - 1)) = 0 And _
                   CInt(strUseValues(Siyou_18.Spacer2 - 1)) = 0 And _
                   CInt(strUseValues(Siyou_18.Spacer3 - 1)) = 0 And _
                   CInt(strUseValues(Siyou_18.Spacer4 - 1)) = 0 Then
                    sbCoordinates.Append(CStr(Siyou_18.Spacer1) & strComma & "0")
                    sbCoordinates.Append(strPipe & CStr(Siyou_18.Spacer2) & strComma & "0")
                    sbCoordinates.Append(strPipe & CStr(Siyou_18.Spacer3) & strComma & "0")
                    sbCoordinates.Append(strPipe & CStr(Siyou_18.Spacer4) & strComma & "0")
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W2280"
                    Exit Function
                End If
            End If

            '7.15 省配線接続仕様・スペーサチェック ～ 7.16 マスキングプレート・スペーサチェック
            For intCI As Integer = 1 To intColR
                '１～９行目が選択されている場合
                For intRI1 As Integer = Siyou_18.Equip To Siyou_18.Elect8
                    '省配線接続仕様・スペーサチェック
                    If Left(strDen.ToString, 1) = "T" And _
                       arySelectInf(intRI1 - 1)(intCI - 1) = "1" And _
                       InStr(strKataValues(intRI1 - 1), "-CL") > 0 Then
                        For intRI2 As Integer = Siyou_18.Spacer1 To Siyou_18.Spacer4
                            If arySelectInf(intRI2 - 1)(intCI - 1) = "1" Then
                                Select Case strKataValues(intRI2 - 1)
                                    Case "4G1-P", "4G1-P-GWS4", "4G1-P-GWS6", "4G2-P", "4G2-P-GWS6", "4G2-P-GWS8", _
                                         "4G1R-P", "4G1R-P-GWS4", "4G1R-P-GWS6", "4G2R-P", "4G2R-P-GWS6", "4G2R-P-GWS8"
                                        sbCoordinates.Append(CStr(intRI2) & strComma & "0")
                                        strMsg = sbCoordinates.ToString
                                        strMsgCd = "W2290"
                                        Exit Function
                                End Select
                            End If
                        Next
                    End If

                    'マスキングプレート・スペーサチェック
                    If arySelectInf(intRI1 - 1)(intCI - 1) = "1" And _
                       InStr(strKataValues(intRI1 - 1), "-MP") > 0 Then
                        For intRI2 As Integer = Siyou_18.Spacer1 To Siyou_18.Spacer4
                            If arySelectInf(intRI2 - 1)(intCI - 1) = "1" Then
                                Select Case strKataValues(intRI2 - 1)
                                    Case "4G1-P", "4G1-P-GWS4", "4G1-P-GWS6", "4G2-P", "4G2-P-GWS6", "4G2-P-GWS8", _
                                         "4G1R-P", "4G1R-P-GWS4", "4G1R-P-GWS6", "4G2R-P", "4G2R-P-GWS6", "4G2R-P-GWS8"
                                        sbCoordinates.Append(CStr(intRI2) & strComma & "0")
                                        strMsg = sbCoordinates.ToString
                                        strMsgCd = "W2300"
                                        Exit Function
                                End Select
                            End If
                        Next
                    End If
                Next
            Next

            fncInputChk = True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try
    End Function
End Class
