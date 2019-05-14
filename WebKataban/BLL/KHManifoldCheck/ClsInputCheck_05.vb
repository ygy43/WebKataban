Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_05

    Public Shared intPosRowCnt As Integer = 23
    Public Shared intColCnt As Integer = 10

    '********************************************************************************************
    '*【関数名】
    '*   fncInpCheck
    '*【処理】
    '*   入力チェック
    '********************************************************************************************
    Public Shared Function fncInpChk(objKtbnStrc As KHKtbnStrc, ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        Dim intElTypeCnt As Integer
        Dim intConBlockCnt As Integer
        Dim intMaxNo As Integer
        Dim intCount As Integer
        Dim sbCoordinates As New StringBuilder
        Dim hshtKataban As New Hashtable
        Dim strCoordinates As String = ""
        Dim bolFlag1 As Boolean
        Dim bolFlag2 As Boolean
        Dim bolFlag3 As Boolean
        Dim strKeyKata As String = objKtbnStrc.strcSelection.strKeyKataban
        strMsg = String.Empty

        fncInpChk = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            '連数設定
            intMaxNo = Int(objKtbnStrc.strcSelection.strOpSymbol(2))

            Select Case strKeyKata
                Case "1"
                    'ミックス時はチェック
                    If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                        '電磁弁チェック
                        '連数までチェックされているか、また重複してチェックしていないかチェックする
                        Dim bolType6 As Boolean = False
                        Dim bolType8 As Boolean = False
                        For intI As Integer = 0 To intMaxNo - 1
                            If arySelectInf(Siyou_05.ElType1 - 1)(intI) = "1" Then
                                bolType6 = True
                            End If
                            If arySelectInf(Siyou_05.ElType2 - 1)(intI) = "1" Then
                                bolType8 = True
                            End If

                            Select Case True
                                Case arySelectInf(Siyou_05.ElType1 - 1)(intI) = "1" And _
                                     arySelectInf(Siyou_05.ElType2 - 1)(intI) = "1"
                                    strMsgCd = "W1790"
                                    Exit Function
                                Case arySelectInf(Siyou_05.ElType1 - 1)(intI) = "1" Or _
                                     arySelectInf(Siyou_05.ElType2 - 1)(intI) = "1"
                                Case Else
                                    strMsgCd = "W1180"
                                    Exit Function
                            End Select
                        Next
                        'PV5-6タイプとPV5-8タイプが両方選択されているかチェックする
                        If Not bolType6 And Not bolType8 Then
                            strMsgCd = "W1790"
                            Exit Function
                        End If

                        '流露遮蔽板のチェック
                        '形番が選択されているかチェック
                        '正しい選択がされているかチェック
                        For intCI As Integer = 0 To intMaxNo - 2
                            For intRI As Integer = Siyou_05.ExpCovRep - 1 To Siyou_05.ExpCovExh - 1
                                If arySelectInf(intRI)(intCI) = "1" Then
                                    Select Case strKataValues(intRI).Trim
                                        Case ""
                                            strMsgCd = "W1400"
                                            Exit Function
                                        Case "CM1-01", "GM1-01" 'GMF Add by Zxjike 2013/10/08
                                            If arySelectInf(Siyou_05.ElType2 - 1)(intCI + 1) = "1" Then
                                                strMsgCd = "W1930"
                                                Exit Function
                                            End If
                                        Case "CM2-01", "GM2-01" 'GMF Add by Zxjike 2013/10/08
                                            If arySelectInf(Siyou_05.ElType1 - 1)(intCI + 1) = "1" Then
                                                strMsgCd = "W1920"
                                                Exit Function
                                            End If
                                    End Select
                                End If
                            Next
                        Next
                    End If
                    'ABポート接続口径
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "HX1" Then
                        For intI As Integer = 0 To intMaxNo - 1
                            If arySelectInf(Siyou_05.ABCon02 - 1)(intI) = "0" And _
                               arySelectInf(Siyou_05.ABCon03 - 1)(intI) = "0" And _
                               arySelectInf(Siyou_05.ABCon04 - 1)(intI) = "0" Then
                                strMsgCd = "W1820"
                                Exit Function
                            End If
                            Select Case arySelectInf(Siyou_05.ABCon02 - 1)(intI)
                                Case "1"
                                    If arySelectInf(Siyou_05.ABCon03 - 1)(intI) = "1" Or _
                                       arySelectInf(Siyou_05.ABCon04 - 1)(intI) = "1" Then
                                        strMsgCd = "W1950"
                                        Exit Function
                                    End If
                            End Select
                            Select Case arySelectInf(Siyou_05.ABCon03 - 1)(intI)
                                Case "1"
                                    If arySelectInf(Siyou_05.ABCon02 - 1)(intI) = "1" Or _
                                       arySelectInf(Siyou_05.ABCon04 - 1)(intI) = "1" Then
                                        strMsgCd = "W1950"
                                        Exit Function
                                    End If
                            End Select
                            Select Case arySelectInf(Siyou_05.ABCon04 - 1)(intI)
                                Case "1"
                                    If arySelectInf(Siyou_05.ABCon02 - 1)(intI) = "1" Or _
                                       arySelectInf(Siyou_05.ABCon03 - 1)(intI) = "1" Then
                                        strMsgCd = "W1950"
                                        Exit Function
                                    End If
                            End Select
                        Next
                    End If
                    'ABポート接続位置
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "L" Then
                        For intI As Integer = 0 To intMaxNo - 1
                            Select Case True
                                Case arySelectInf(Siyou_05.ABPlugR - 1)(intI) = "1" And _
                                     arySelectInf(Siyou_05.ABPlugL - 1)(intI) = "1"
                                    strMsgCd = "W1940"
                                    Exit Function
                                Case arySelectInf(Siyou_05.ABPlugR - 1)(intI) = "1" Or _
                                     arySelectInf(Siyou_05.ABPlugL - 1)(intI) = "1"
                                Case Else
                                    strMsgCd = "W1810"
                                    Exit Function
                            End Select
                        Next
                    End If
                    fncInpChk = True
                Case "4", "6"
                    sbCoordinates = New StringBuilder

                    'Ａ・Ｂポートプラグ位置チェック
                    If objKtbnStrc.strcSelection.strOpSymbol(4) = "L" Then
                        'Ａ・Ｂポートプラグ位置が一つも選択されていない場合、エラー
                        If Int(strUseValues(Siyou_05.ABPlugR - 1)) = 0 And _
                           Int(strUseValues(Siyou_05.ABPlugL - 1)) = 0 Then
                            strCoordinates = Siyou_05.ABPlugR & strComma & "0" & strPipe & _
                                             Siyou_05.ABPlugL & strComma & "0"
                            strMsg = strCoordinates
                            strMsgCd = "W1800"
                            Exit Function
                        End If
                        For intCI As Integer = 2 To intMaxNo - 1
                            Select Case True
                                Case arySelectInf(Siyou_05.ABPlugR - 1)(intCI) = "0" And _
                                     arySelectInf(Siyou_05.ABPlugL - 1)(intCI) = "0"
                                    '選択されていない列がある場合、エラー
                                    strCoordinates = Siyou_05.ABPlugR & strComma & CStr(intCI + 1) & strPipe & _
                                                     Siyou_05.ABPlugL & strComma & CStr(intCI + 1)
                                    strMsg = strCoordinates
                                    strMsgCd = "W1810"
                                    Exit Function
                                Case arySelectInf(Siyou_05.ABPlugR - 1)(intCI) = "1" And _
                                     arySelectInf(Siyou_05.ABPlugL - 1)(intCI) = "1"
                                    '一列につき２つ以上選択されていたらエラー
                                    strCoordinates = Siyou_05.ABPlugR & strComma & CStr(intCI + 1) & strPipe & _
                                                     Siyou_05.ABPlugL & strComma & CStr(intCI + 1)
                                    strMsg = strCoordinates
                                    strMsgCd = "W1940"
                                    Exit Function
                            End Select
                        Next
                    End If
                    'Ａ・Ｂポート接続口径チェック
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "HX1"
                            intCount = 0
                            '選択状態の形番数をカウント
                            For intRI As Integer = Siyou_05.ABCon02 - 1 To Siyou_05.ABCon04 - 1
                                If Int(strUseValues(intRI)) > 0 Then
                                    intCount = intCount + 1
                                End If
                            Next
                            'カウントが０または１の場合、エラー
                            Select Case intCount
                                Case 0
                                    sbCoordinates.Append(Siyou_05.ABCon02 & strComma & "0" & strPipe)
                                    sbCoordinates.Append(Siyou_05.ABCon03 & strComma & "0" & strPipe)
                                    sbCoordinates.Append(Siyou_05.ABCon04 & strComma & "0")
                                    strMsg = strCoordinates.ToString
                                    strMsgCd = "W1820"
                                    Exit Function
                                Case 1
                                    sbCoordinates.Append(Siyou_05.ABCon02 & strComma & "0" & strPipe)
                                    sbCoordinates.Append(Siyou_05.ABCon03 & strComma & "0" & strPipe)
                                    sbCoordinates.Append(Siyou_05.ABCon04 & strComma & "0")
                                    strMsg = strCoordinates.ToString
                                    strMsgCd = "W1830"
                                    Exit Function
                            End Select
                    End Select
                    '一列につき２つ以上選択されていたらエラー
                    For intCI As Integer = 2 To intMaxNo - 1
                        intCount = 0
                        sbCoordinates = New StringBuilder
                        For intRI As Integer = Siyou_05.ABCon02 - 1 To Siyou_05.ABCon04 - 1
                            If arySelectInf(intRI)(intCI) = "1" Then
                                intCount = intCount + 1
                                sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                            End If
                        Next
                        If intCount > 1 Then
                            strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                            strMsg = strCoordinates.ToString
                            strMsgCd = "W1950"
                            Exit Function
                        End If
                    Next
                    fncInpChk = True
                Case "8", "9"
                    If objKtbnStrc.strcSelection.strOpSymbol(9) = "" Then
                        sbCoordinates = New StringBuilder

                        'Ａ・Ｂポートプラグ位置チェック
                        If objKtbnStrc.strcSelection.strOpSymbol(4) = "L" Then
                            'Ａ・Ｂポートプラグ位置が一つも選択されていない場合、エラー
                            If Int(strUseValues(Siyou_05.ABPlugR - 1)) = 0 And Int(strUseValues(Siyou_05.ABPlugL - 1)) = 0 Then
                                strCoordinates = Siyou_05.ABPlugR & strComma & "0" & strPipe & Siyou_05.ABPlugL & strComma & "0"
                                strMsg = strCoordinates.ToString
                                strMsgCd = "W1800"
                                Exit Function
                            End If
                            For intCI As Integer = 0 To intMaxNo - 1
                                Select Case True
                                    Case arySelectInf(Siyou_05.ABPlugR - 1)(intCI) = "0" And _
                                         arySelectInf(Siyou_05.ABPlugL - 1)(intCI) = "0"
                                        '選択されていない列がある場合、エラー
                                        strCoordinates = Siyou_05.ABPlugR & strComma & CStr(intCI + 1) & strPipe & _
                                                         Siyou_05.ABPlugL & strComma & CStr(intCI + 1)
                                        strMsg = strCoordinates.ToString
                                        strMsgCd = "W1810"
                                        Exit Function
                                    Case arySelectInf(Siyou_05.ABPlugR - 1)(intCI) = "1" And _
                                         arySelectInf(Siyou_05.ABPlugL - 1)(intCI) = "1"
                                        '一列につき２つ以上選択されていたらエラー
                                        strCoordinates = Siyou_05.ABPlugR & strComma & CStr(intCI + 1) & strPipe & _
                                                         Siyou_05.ABPlugL & strComma & CStr(intCI + 1)
                                        strMsg = strCoordinates.ToString
                                        strMsgCd = "W1940"
                                        Exit Function
                                End Select
                            Next
                        End If
                        'Ａ・Ｂポート接続口径チェック
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Case "HX1"
                                intCount = 0
                                '選択状態の形番数をカウント
                                For intRI As Integer = Siyou_05.ABCon02 - 1 To Siyou_05.ABCon04 - 1
                                    If Int(strUseValues(intRI)) > 0 Then
                                        intCount = intCount + 1
                                    End If
                                Next
                                'カウントが０または１の場合、エラー
                                Select Case intCount
                                    Case 0
                                        sbCoordinates.Append(Siyou_05.ABCon02 & strComma & "0" & strPipe)
                                        sbCoordinates.Append(Siyou_05.ABCon03 & strComma & "0" & strPipe)
                                        sbCoordinates.Append(Siyou_05.ABCon04 & strComma & "0")
                                        strMsg = strCoordinates.ToString
                                        strMsgCd = "W1820"
                                        Exit Function
                                    Case 1
                                        sbCoordinates.Append(Siyou_05.ABCon02 & strComma & "0" & strPipe)
                                        sbCoordinates.Append(Siyou_05.ABCon03 & strComma & "0" & strPipe)
                                        sbCoordinates.Append(Siyou_05.ABCon04 & strComma & "0")
                                        strMsg = strCoordinates.ToString
                                        strMsgCd = "W1830"
                                        Exit Function
                                End Select
                        End Select
                        '一列につき２つ以上選択されていたらエラー
                        For intCI As Integer = 0 To intMaxNo - 1
                            intCount = 0
                            sbCoordinates = New StringBuilder
                            For intRI As Integer = Siyou_05.ABCon02 - 1 To Siyou_05.ABCon04 - 1
                                If arySelectInf(intRI)(intCI) = "1" Then
                                    intCount = intCount + 1
                                    sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                                End If
                            Next
                            If intCount > 1 Then
                                strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                                strMsg = strCoordinates.ToString
                                strMsgCd = "W1950"
                                Exit Function
                            End If
                        Next
                        fncInpChk = True
                    End If
            End Select

            'DELETE BY YGY 20141118
            'If Not fncInpChk Then
            Select Case strKeyKata
                Case "5", "7", "9"
                    '************** 形番チェック ***********************************************
                    For intRI As Integer = 2 To Siyou_05.ExpCovExh - 1
                        '形番が未選択の場合
                        If Len(Trim(strKataValues(intRI))) = 0 Then
                            'ABポート以外の行で設置位置が選択されていたらエラー
                            If (intRI < Siyou_05.ABPlugR - 1 Or _
                                intRI > Siyou_05.ABCon04 - 1) And _
                                Int(strUseValues(intRI)) > 0 Then
                                strMsgCd = "W1400"
                                Exit Function
                            End If
                        ElseIf intRI < Siyou_05.ExpCovRep - 1 Then
                            '形番重複チェック
                            If hshtKataban.ContainsKey(strKataValues(intRI)) Then
                                strMsgCd = "W1330"
                                Exit Function
                            Else
                                hshtKataban.Add(strKataValues(intRI), "")
                            End If
                        End If
                    Next

                    '************** 電磁連数弁形式チェック(1.1) ************************************
                    '一列ごとの電磁弁形式選択数が１以外の場合、エラー
                    For intCI As Integer = 2 To intMaxNo - 1
                        intElTypeCnt = 0
                        sbCoordinates = New StringBuilder
                        For intRI As Integer = Siyou_05.ElType1 - 1 To Siyou_05.ElType6 - 1
                            If arySelectInf(intRI)(intCI) = "1" Then
                                sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                                intElTypeCnt = intElTypeCnt + 1
                            End If
                        Next
                        If intElTypeCnt <> 1 Then
                            If Len(sbCoordinates.ToString) = 0 Then
                                For intRI As Integer = Siyou_05.ElType1 - 1 To Siyou_05.ElType6 - 1
                                    sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                                Next
                            End If
                            strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                            strMsg = strCoordinates.ToString
                            strMsgCd = "W1790"
                            Exit Function
                        End If
                    Next

                    sbCoordinates = New StringBuilder
                    bolFlag1 = False
                    bolFlag2 = False

                    '接続ブロックカウント取得
                    intConBlockCnt = fncGetConectBlockCnt(objKtbnStrc, sbCoordinates, intMaxNo)

                    If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                        'ミックス指定の場合、電磁弁形式で両方のタイプを選択していないとエラー
                        For intRI As Integer = Siyou_05.ElType1 - 1 To Siyou_05.ElType6 - 1
                            If Int(strUseValues(intRI)) > 0 Then
                                If Left(strKataValues(intRI), 3) = "PV5" Then
                                    If strKataValues(intRI).Contains("6") Then
                                        bolFlag1 = True
                                    ElseIf strKataValues(intRI).Contains("8") Then
                                        bolFlag2 = True
                                    End If
                                End If
                            End If
                        Next
                        If Not (bolFlag1 And bolFlag2) Then
                            strMsgCd = "W1790"
                            Exit Function
                        End If

                        'ミックス指定の場合、接続ブロッカウントが２以上だったらエラー
                        If intConBlockCnt > 1 Then
                            strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                            strMsg = strCoordinates.ToString
                            strMsgCd = "W1790"
                            Exit Function
                        End If
                    End If
                    sbCoordinates = New StringBuilder

                    '************** Ａ・Ｂポートプラグ位置チェック(1.2) ************************************
                    If objKtbnStrc.strcSelection.strOpSymbol(4) = "L" Then
                        'Ａ・Ｂポートプラグ位置が一つも選択されていない場合、エラー
                        If Int(strUseValues(Siyou_05.ABPlugR - 1)) = 0 And Int(strUseValues(Siyou_05.ABPlugL - 1)) = 0 Then

                            strCoordinates = Siyou_05.ABPlugR & strComma & "0" & strPipe & Siyou_05.ABPlugL & strComma & "0"
                            strMsg = strCoordinates.ToString
                            strMsgCd = "W1800"
                            Exit Function
                        End If

                        '選択されていない列がある場合、エラー
                        For intCI As Integer = 2 To intMaxNo - 1
                            If arySelectInf(Siyou_05.ABPlugR - 1)(intCI) = "0" And arySelectInf(Siyou_05.ABPlugL - 1)(intCI) = "0" Then

                                strCoordinates = Siyou_05.ABPlugR & strComma & CStr(intCI + 1) & strPipe & _
                                                 Siyou_05.ABPlugL & strComma & CStr(intCI + 1)
                                strMsg = strCoordinates.ToString
                                strMsgCd = "W1810"
                                Exit Function
                            End If
                        Next
                    End If
                    '一列につき２つ以上選択されていたらエラー
                    For intCI As Integer = 2 To intMaxNo - 1
                        If arySelectInf(Siyou_05.ABPlugR - 1)(intCI) = "1" And arySelectInf(Siyou_05.ABPlugL - 1)(intCI) = "1" Then
                            strCoordinates = Siyou_05.ABPlugR & strComma & CStr(intCI + 1) & strPipe & _
                                             Siyou_05.ABPlugL & strComma & CStr(intCI + 1)
                            strMsg = strCoordinates.ToString
                            strMsgCd = "W1940"
                            Exit Function
                        End If
                    Next

                    '************** Ａ・Ｂポート接続口径チェック(1.3) ************************************
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "HX1", "HX2"
                            intCount = 0

                            '選択状態の形番数をカウント
                            For intRI As Integer = Siyou_05.ABCon02 - 1 To Siyou_05.ABCon04 - 1
                                If Int(strUseValues(intRI)) > 0 Then
                                    intCount = intCount + 1
                                End If
                            Next

                            'カウントが０または１の場合、エラー
                            If intCount = 0 Then
                                sbCoordinates.Append(Siyou_05.ABCon02 & strComma & "0" & strPipe)
                                sbCoordinates.Append(Siyou_05.ABCon03 & strComma & "0" & strPipe)
                                sbCoordinates.Append(Siyou_05.ABCon04 & strComma & "0")
                                strMsg = strCoordinates.ToString
                                strMsgCd = "W1820"
                                Exit Function

                            ElseIf intCount = 1 Then
                                sbCoordinates.Append(Siyou_05.ABCon02 & strComma & "0" & strPipe)
                                sbCoordinates.Append(Siyou_05.ABCon03 & strComma & "0" & strPipe)
                                sbCoordinates.Append(Siyou_05.ABCon04 & strComma & "0")
                                strMsg = strCoordinates.ToString
                                strMsgCd = "W1830"
                                Exit Function
                            End If
                    End Select
                    '一列につき２つ以上選択されていたらエラー
                    For intCI As Integer = 2 To intMaxNo - 1
                        intCount = 0
                        sbCoordinates = New StringBuilder
                        For intRI As Integer = Siyou_05.ABCon02 - 1 To Siyou_05.ABCon04 - 1
                            If arySelectInf(intRI)(intCI) = "1" Then
                                intCount = intCount + 1
                                sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                            End If
                        Next
                        If intCount > 1 Then
                            strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                            strMsg = strCoordinates.ToString
                            strMsgCd = "W1950"
                            Exit Function
                        End If
                    Next
                    sbCoordinates = New StringBuilder

                    '************** 給気スペーサチェック(1.4) ************************************
                    For intCI As Integer = 2 To Int(objKtbnStrc.strcSelection.strOpSymbol(2)) - 1
                        intCount = 0
                        For intRI As Integer = Siyou_05.RepSpace1 - 1 To Siyou_05.RepSpace2 - 1
                            If arySelectInf(intRI)(intCI) = "1" Then

                                intCount = intCount + 1
                                sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                                If intCount > 1 Then
                                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                                    strMsg = strCoordinates.ToString
                                    strMsgCd = "W1960"
                                    Exit Function
                                End If

                                If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                                    If Left(strKataValues(intRI), 7) = "CMF1-P-" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "6", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1850"
                                            Exit Function
                                        End If
                                    ElseIf Left(strKataValues(intRI), 7) = "CMF2-P-" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "8", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1840"
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        sbCoordinates = New StringBuilder
                    Next

                    '************** 排気スペーサチェック(1.5) ************************************
                    For intCI As Integer = 2 To Int(objKtbnStrc.strcSelection.strOpSymbol(2)) - 1
                        intCount = 0
                        For intRI As Integer = Siyou_05.ExhSpace1 - 1 To Siyou_05.ExhSpace2 - 1
                            If arySelectInf(intRI)(intCI) = "1" Then

                                intCount = intCount + 1
                                sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                                If intCount > 1 Then
                                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                                    strMsg = strCoordinates.ToString
                                    strMsgCd = "W1970"
                                    Exit Function
                                End If

                                If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                                    If Left(strKataValues(intRI), 7) = "CMF1-R-" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "6", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1870"
                                            Exit Function
                                        End If
                                    ElseIf Left(strKataValues(intRI), 7) = "CMF2-R-" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "8", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1860"
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        sbCoordinates = New StringBuilder
                    Next

                    '************** パイロット弁チェック(1.6) ************************************
                    For intCI As Integer = 2 To Int(objKtbnStrc.strcSelection.strOpSymbol(2)) - 1
                        intCount = 0
                        For intRI As Integer = Siyou_05.Pilot1 - 1 To Siyou_05.Pilot2 - 1
                            If arySelectInf(intRI)(intCI) = "1" Then

                                intCount = intCount + 1
                                sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                                If intCount > 1 Then
                                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                                    strMsg = strCoordinates.ToString
                                    strMsgCd = "W1980"
                                    Exit Function
                                End If

                                If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                                    If Left(strKataValues(intRI), 7) = "CMF1-PC" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "6", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1890"
                                            Exit Function
                                        End If
                                    ElseIf Left(strKataValues(intRI), 7) = "CMF2-PC" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "8", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1880"
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        sbCoordinates = New StringBuilder
                    Next

                    '************** スペーサ形減圧弁チェック(1.7) ************************************
                    For intCI As Integer = 2 To Int(objKtbnStrc.strcSelection.strOpSymbol(2)) - 1
                        intCount = 0
                        For intRI As Integer = Siyou_05.SpDecomp1 - 1 To Siyou_05.SpDecomp4 - 1
                            If arySelectInf(intRI)(intCI) = "1" Then

                                intCount = intCount + 1
                                sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                                If intCount > 1 Then
                                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                                    strMsg = strCoordinates.ToString
                                    strMsgCd = "W1990"
                                    Exit Function
                                End If

                                If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                                    If Left(strKataValues(intRI), 8) = "CMF1-SR-" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "6", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1910"
                                            Exit Function
                                        End If
                                    ElseIf Left(strKataValues(intRI), 8) = "CMF2-SR-" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "8", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1900"
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        sbCoordinates = New StringBuilder
                    Next

                    '************** 流露遮蔽板チェック(1.8) ************************************
                    If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                        For intCI As Integer = 2 To intColCnt - 2
                            For intRI As Integer = Siyou_05.ExpCovRep - 1 To Siyou_05.ExpCovExh - 1
                                If arySelectInf(intRI)(intCI) = "1" Then

                                    'Add by Zxjike 2013/10/08
                                    Select Case Left(strKataValues(intRI), 6)
                                        Case "CM1-01", "GM1-01"
                                            If Not fncCheckElType(objKtbnStrc, intCI, "6", strCoordinates) Then
                                                strMsg = strCoordinates.ToString
                                                strMsgCd = "W1930"
                                                Exit Function
                                            End If
                                        Case "CM2-01", "GM2-01"
                                            If Not fncCheckElType(objKtbnStrc, intCI, "8", strCoordinates) Then
                                                strMsg = strCoordinates.ToString
                                                strMsgCd = "W1920"
                                                Exit Function
                                            End If
                                    End Select
                                    'Del by Zxjike 2013/10/08
                                    'If Left(strKataValues(intRI), 6) = "CM1-01" Then
                                    '    If Not fncCheckElType(intCI, "6", strKataValues, arySelectInf, strCoordinates, intRI + 1) Then
                                    '        strMsg = strCoordinates.ToString
                                    '        strMsgCd = "W1930"
                                    '        Exit Function
                                    '    End If
                                    'ElseIf Left(strKataValues(intRI), 6) = "CM2-01" Then
                                    '    If Not fncCheckElType(intCI, "8", strKataValues, arySelectInf, strCoordinates, intRI + 1) Then
                                    '        strMsg = strCoordinates.ToString
                                    '        strMsgCd = "W1920"
                                    '        Exit Function
                                    '    End If
                                    'End If
                                End If
                            Next
                        Next
                    End If
                Case Else
                    '************** 形番チェック ***********************************************
                    For intRI As Integer = 0 To Siyou_05.ExpCovExh - 1
                        '形番が未選択の場合
                        If Len(Trim(strKataValues(intRI))) = 0 Then
                            'ABポート以外の行で設置位置が選択されていたらエラー
                            If (intRI < Siyou_05.ABPlugR - 1 Or _
                                intRI > Siyou_05.ABCon04 - 1) And _
                                Int(strUseValues(intRI)) > 0 Then
                                strMsgCd = "W1400"
                                Exit Function
                            End If
                        ElseIf intRI < Siyou_05.ExpCovRep - 1 Then
                            '形番重複チェック
                            If hshtKataban.ContainsKey(strKataValues(intRI)) Then
                                strMsgCd = "W1330"
                                Exit Function
                            Else
                                hshtKataban.Add(strKataValues(intRI), "")
                            End If
                        End If
                    Next

                    '************** 電磁連数弁形式チェック(1.1) ************************************
                    '一列ごとの電磁弁形式選択数が１以外の場合、エラー
                    For intCI As Integer = 0 To intMaxNo - 1
                        intElTypeCnt = 0
                        sbCoordinates = New StringBuilder
                        For intRI As Integer = Siyou_05.ElType1 - 1 To Siyou_05.ElType6 - 1
                            If arySelectInf(intRI)(intCI) = "1" Then
                                sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                                intElTypeCnt = intElTypeCnt + 1
                            End If
                        Next
                        If intElTypeCnt <> 1 Then
                            If Len(sbCoordinates.ToString) = 0 Then
                                For intRI As Integer = Siyou_05.ElType1 - 1 To Siyou_05.ElType6 - 1
                                    sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                                Next
                            End If
                            strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                            strMsg = strCoordinates.ToString
                            strMsgCd = "W1790"
                            Exit Function
                        End If
                    Next

                    sbCoordinates = New StringBuilder
                    bolFlag1 = False
                    bolFlag2 = False

                    '接続ブロックカウント取得
                    intConBlockCnt = fncGetConectBlockCnt(objKtbnStrc, sbCoordinates, intMaxNo)

                    If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                        'ミックス指定の場合、電磁弁形式で両方のタイプを選択していないとエラー
                        For intRI As Integer = Siyou_05.ElType1 - 1 To Siyou_05.ElType6 - 1
                            If Int(strUseValues(intRI)) > 0 Then
                                'If Left(strKataValues(intRI), 3) = "PV5" Then
                                '    If strKataValues(intRI).Contains("6") Then
                                '        bolFlag1 = True
                                '    ElseIf strKataValues(intRI).Contains("8") Then
                                '        bolFlag2 = True
                                '    End If
                                'End If
                                'Add by Zxjike 2013/10/08 ↓
                                If strKataValues(intRI).StartsWith("PV5-6") Or strKataValues(intRI).StartsWith("PV5G-6") Then
                                    bolFlag1 = True
                                End If
                                If strKataValues(intRI).StartsWith("PV5-8") Or strKataValues(intRI).StartsWith("PV5G-8") Then
                                    bolFlag2 = True
                                End If
                                'Add by Zxjike 2013/10/08 ↑
                                If Left(strKataValues(intRI), 2) = "CM" Then
                                    If strKataValues(intRI).Contains("1") Then
                                        bolFlag1 = True
                                    ElseIf strKataValues(intRI).Contains("2") Then
                                        bolFlag2 = True
                                    End If
                                End If
                            End If
                        Next
                        If Not (bolFlag1 = True And bolFlag2 = True) Then
                            strMsgCd = "W1790"
                            Exit Function
                        End If

                        'ミックス指定の場合、接続ブロッカウントが２以上だったらエラー
                        If intConBlockCnt > 1 Then
                            strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                            strMsg = strCoordinates.ToString
                            strMsgCd = "W1790"
                            Exit Function
                        End If
                    End If

                    bolFlag1 = False
                    bolFlag2 = False
                    bolFlag3 = False

                    '電磁弁組合せチェック
                    '[YZ-S][YZ-D]を含む電磁弁と、その他の電磁弁は同じマニホールドに組み付けることは出来ない
                    For intCI As Integer = 0 To intMaxNo - 1
                        For intRI As Integer = Siyou_05.ElType1 - 1 To Siyou_05.ElType6 - 1
                            If arySelectInf(intRI)(intCI) = "1" Then
                                If InStr(strKataValues(intRI), "YZ-S") Or InStr(strKataValues(intRI), "YZ-D") Then
                                    bolFlag1 = True
                                Else
                                    bolFlag2 = True
                                End If
                            End If
                        Next
                        For intRI As Integer = Siyou_05.RepSpace1 - 1 To Siyou_05.Pilot2 - 1
                            If arySelectInf(intRI)(intCI) = "1" Then
                                bolFlag3 = True
                            End If
                        Next
                    Next
                    If bolFlag1 = True And bolFlag2 = True Then
                        strMsgCd = "W8620"
                        Exit Function
                    End If
                    If bolFlag1 = True And bolFlag3 = True Then
                        strMsgCd = "W8630"
                        Exit Function
                    End If

                    sbCoordinates = New StringBuilder

                    '************** Ａ・Ｂポートプラグ位置チェック(1.2) ************************************
                    If objKtbnStrc.strcSelection.strOpSymbol(4) = "L" Then
                        'Ａ・Ｂポートプラグ位置が一つも選択されていない場合、エラー
                        If Int(strUseValues(Siyou_05.ABPlugR - 1)) = 0 And Int(strUseValues(Siyou_05.ABPlugL - 1)) = 0 Then
                            strCoordinates = Siyou_05.ABPlugR & strComma & "0" & strPipe & Siyou_05.ABPlugL & strComma & "0"
                            strMsg = strCoordinates.ToString
                            strMsgCd = "W1800"
                            Exit Function
                        End If

                        '選択されていない列がある場合、エラー
                        For intCI As Integer = 0 To intMaxNo - 1
                            If arySelectInf(Siyou_05.ABPlugR - 1)(intCI) = "0" And arySelectInf(Siyou_05.ABPlugL - 1)(intCI) = "0" Then
                                strCoordinates = Siyou_05.ABPlugR & strComma & CStr(intCI + 1) & strPipe & _
                                                 Siyou_05.ABPlugL & strComma & CStr(intCI + 1)
                                strMsg = strCoordinates.ToString
                                strMsgCd = "W1810"
                                Exit Function
                            End If
                        Next
                    End If

                    '一列につき２つ以上選択されていたらエラー
                    For intCI As Integer = 0 To intMaxNo - 1
                        If arySelectInf(Siyou_05.ABPlugR - 1)(intCI) = "1" And arySelectInf(Siyou_05.ABPlugL - 1)(intCI) = "1" Then
                            strCoordinates = Siyou_05.ABPlugR & strComma & CStr(intCI + 1) & strPipe & _
                                             Siyou_05.ABPlugL & strComma & CStr(intCI + 1)
                            strMsg = strCoordinates.ToString
                            strMsgCd = "W1940"
                            Exit Function
                        End If
                    Next

                    '************** Ａ・Ｂポート接続口径チェック(1.3) ************************************
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "HX1", "HX2"
                            intCount = 0

                            '選択状態の形番数をカウント
                            For intRI As Integer = Siyou_05.ABCon02 - 1 To Siyou_05.ABCon04 - 1
                                If Int(strUseValues(intRI)) > 0 Then
                                    intCount = intCount + 1
                                End If
                            Next

                            'カウントが０または１の場合、エラー
                            If intCount = 0 Then
                                sbCoordinates.Append(Siyou_05.ABCon02 & strComma & "0" & strPipe)
                                sbCoordinates.Append(Siyou_05.ABCon03 & strComma & "0" & strPipe)
                                sbCoordinates.Append(Siyou_05.ABCon04 & strComma & "0")
                                strMsg = strCoordinates.ToString
                                strMsgCd = "W1820"
                                Exit Function
                            ElseIf intCount = 1 Then
                                sbCoordinates.Append(Siyou_05.ABCon02 & strComma & "0" & strPipe)
                                sbCoordinates.Append(Siyou_05.ABCon03 & strComma & "0" & strPipe)
                                sbCoordinates.Append(Siyou_05.ABCon04 & strComma & "0")
                                strMsg = strCoordinates.ToString
                                strMsgCd = "W1830"
                                Exit Function
                            End If
                    End Select
                    '一列につき２つ以上選択されていたらエラー
                    For intCI As Integer = 0 To intMaxNo - 1
                        intCount = 0
                        sbCoordinates = New StringBuilder
                        For intRI As Integer = Siyou_05.ABCon02 - 1 To Siyou_05.ABCon04 - 1
                            If arySelectInf(intRI)(intCI) = "1" Then
                                intCount = intCount + 1
                                sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                            End If
                        Next
                        If intCount > 1 Then
                            strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                            strMsg = strCoordinates.ToString
                            strMsgCd = "W1950"
                            Exit Function
                        End If
                    Next
                    sbCoordinates = New StringBuilder

                    '************** 給気スペーサチェック(1.4) ************************************
                    For intCI As Integer = 0 To Int(objKtbnStrc.strcSelection.strOpSymbol(2)) - 1
                        intCount = 0
                        For intRI As Integer = Siyou_05.RepSpace1 - 1 To Siyou_05.RepSpace2 - 1
                            If arySelectInf(intRI)(intCI) = "1" Then

                                intCount = intCount + 1
                                sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                                If intCount > 1 Then
                                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                                    strMsg = strCoordinates.ToString
                                    strMsgCd = "W1960"
                                    Exit Function
                                End If

                                If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                                    'Add by Zxjike 2013/10/08
                                    If Not fncCheckMP(objKtbnStrc, intCI, strCoordinates) Then
                                        strMsg = strCoordinates.ToString
                                        strMsgCd = "W8990"
                                        Exit Function
                                    End If
                                    If Left(strKataValues(intRI), 7) = "CMF1-P-" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "6", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1850"
                                            Exit Function
                                        End If
                                    ElseIf Left(strKataValues(intRI), 7) = "CMF2-P-" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "8", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1840"
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        sbCoordinates = New StringBuilder
                    Next

                    '************** 排気スペーサチェック(1.5) ************************************
                    For intCI As Integer = 0 To Int(objKtbnStrc.strcSelection.strOpSymbol(2)) - 1
                        intCount = 0
                        For intRI As Integer = Siyou_05.ExhSpace1 - 1 To Siyou_05.ExhSpace2 - 1
                            If arySelectInf(intRI)(intCI) = "1" Then

                                intCount = intCount + 1
                                sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                                If intCount > 1 Then
                                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                                    strMsg = strCoordinates.ToString
                                    strMsgCd = "W1970"
                                    Exit Function
                                End If

                                If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                                    'Add by Zxjike 2013/10/08
                                    If Not fncCheckMP(objKtbnStrc, intCI, strCoordinates) Then
                                        strMsg = strCoordinates.ToString
                                        strMsgCd = "W8990"
                                        Exit Function
                                    End If
                                    If Left(strKataValues(intRI), 7) = "CMF1-R-" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "6", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1870"
                                            Exit Function
                                        End If
                                    ElseIf Left(strKataValues(intRI), 7) = "CMF2-R-" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "8", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1860"
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        sbCoordinates = New StringBuilder
                    Next

                    '************** パイロット弁チェック(1.6) ************************************
                    For intCI As Integer = 0 To Int(objKtbnStrc.strcSelection.strOpSymbol(2)) - 1
                        intCount = 0
                        For intRI As Integer = Siyou_05.Pilot1 - 1 To Siyou_05.Pilot2 - 1
                            If arySelectInf(intRI)(intCI) = "1" Then
                                intCount = intCount + 1
                                sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                                If intCount > 1 Then
                                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                                    strMsg = strCoordinates.ToString
                                    strMsgCd = "W1980"
                                    Exit Function
                                End If

                                If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                                    'Add by Zxjike 2013/10/08
                                    If Not fncCheckMP(objKtbnStrc, intCI, strCoordinates) Then
                                        strMsg = strCoordinates.ToString
                                        strMsgCd = "W8990"
                                        Exit Function
                                    End If
                                    If Left(strKataValues(intRI), 7) = "CMF1-PC" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "6", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1890"
                                            Exit Function
                                        End If
                                    ElseIf Left(strKataValues(intRI), 7) = "CMF2-PC" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "8", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1880"
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        sbCoordinates = New StringBuilder
                    Next

                    '************** スペーサ形減圧弁チェック(1.7) ************************************
                    For intCI As Integer = 0 To Int(objKtbnStrc.strcSelection.strOpSymbol(2)) - 1
                        intCount = 0
                        For intRI As Integer = Siyou_05.SpDecomp1 - 1 To Siyou_05.SpDecomp4 - 1
                            If arySelectInf(intRI)(intCI) = "1" Then

                                intCount = intCount + 1
                                sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                                If intCount > 1 Then
                                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                                    strMsg = strCoordinates.ToString
                                    strMsgCd = "W1990"
                                    Exit Function
                                End If

                                If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                                    'Add by Zxjike 2013/10/08
                                    If Not fncCheckMP(objKtbnStrc, intCI, strCoordinates) Then
                                        strMsg = strCoordinates.ToString
                                        strMsgCd = "W8990"
                                        Exit Function
                                    End If
                                    If Left(strKataValues(intRI), 8) = "CMF1-SR-" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "6", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1910"
                                            Exit Function
                                        End If
                                    ElseIf Left(strKataValues(intRI), 8) = "CMF2-SR-" Then
                                        If Not fncCheckElType(objKtbnStrc, intCI, "8", strCoordinates, intRI + 1) Then
                                            strMsg = strCoordinates.ToString
                                            strMsgCd = "W1900"
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        sbCoordinates = New StringBuilder
                    Next

                    '************** 流露遮蔽板チェック(1.8) ************************************
                    If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                        For intCI As Integer = 0 To intColCnt - 2
                            For intRI As Integer = Siyou_05.ExpCovRep - 1 To Siyou_05.ExpCovExh - 1
                                If arySelectInf(intRI)(intCI) = "1" Then

                                    'Add by Zxjike 2013/10/08
                                    Select Case Left(strKataValues(intRI), 6)
                                        Case "CM1-01", "GM1-01"
                                            If Not fncCheckElType(objKtbnStrc, intCI, "6", strCoordinates) Then
                                                strMsg = strCoordinates.ToString
                                                strMsgCd = "W1930"
                                                Exit Function
                                            End If
                                        Case "CM2-01", "GM2-01"
                                            If Not fncCheckElType(objKtbnStrc, intCI, "8", strCoordinates) Then
                                                strMsg = strCoordinates.ToString
                                                strMsgCd = "W1920"
                                                Exit Function
                                            End If
                                    End Select
                                    'Del by Zxjike 2013/10/08
                                    'If Left(strKataValues(intRI), 6) = "CM1-01" Then
                                    '    If Not fncCheckElType(intCI, "6", strKataValues, arySelectInf, strCoordinates, intRI + 1) Then
                                    '        strMsg = strCoordinates.ToString
                                    '        strMsgCd = "W1930"
                                    '        Exit Function
                                    '    End If
                                    'ElseIf Left(strKataValues(intRI), 6) = "CM2-01" Then
                                    '    If Not fncCheckElType(intCI, "8", strKataValues, arySelectInf, strCoordinates, intRI + 1) Then
                                    '        strMsg = strCoordinates.ToString
                                    '        strMsgCd = "W1920"
                                    '        Exit Function
                                    '    End If
                                    'End If
                                End If
                            Next
                        Next
                    End If
            End Select

            '選択オプション数チェック
            intCount = 0
            For intRI As Integer = 0 To Siyou_05.ExpCovExh - 1
                If Int(strUseValues(intRI)) > 0 Then
                    intCount = intCount + 1
                End If
            Next
            If intConBlockCnt > 0 Then
                intCount = intCount + 1
            End If
            If intCount > 20 Then
                strMsgCd = "W2000"
                Exit Function
            End If
            fncInpChk = True
            'End If
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncGetConectBlockCnt
    '*【処理】
    '*   接続ブロック数をカウントする
    '********************************************************************************************
    Public Shared Function fncGetConectBlockCnt(objKtbnStrc As KHKtbnStrc, ByRef sbCoordinates As StringBuilder, _
                                                ByVal intRensuu As Integer) As Integer

        Dim intConBlockCnt As Integer = 0
        Dim strSaveType() As String

        Try
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            ReDim strSaveType(2)
            For intCI As Integer = 0 To Int(intRensuu) - 1
                For intRI As Integer = Siyou_05.ElType1 - 1 To Siyou_05.ElType6 - 1
                    If objKtbnStrc.strcSelection.strPositionInfo(intRI)(intCI) = "1" Then
                        If strSaveType(0) Is Nothing Then
                        ElseIf Not (strKataValues(intRI).Contains(strSaveType(0)) Or _
                                    strKataValues(intRI).Contains(strSaveType(1)) Or _
                                    strKataValues(intRI).Contains(strSaveType(2))) Then
                            intConBlockCnt = intConBlockCnt + 1
                            sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                        End If
                        'Add by Zxjike 2013/10/08
                        If strKataValues(intRI).StartsWith("PV5-6") Or _
                            strKataValues(intRI).StartsWith("PV5G-6") Or _
                            strKataValues(intRI).StartsWith("CM1") Then
                            strSaveType(0) = "PV5-6"
                            strSaveType(1) = "PV5G-6"
                            strSaveType(2) = "CM1"
                        End If
                        If strKataValues(intRI).StartsWith("PV5-8") Or _
                            strKataValues(intRI).StartsWith("PV5G-8") Or _
                            strKataValues(intRI).StartsWith("CM2") Then
                            strSaveType(0) = "PV5-8"
                            strSaveType(1) = "PV5G-8"
                            strSaveType(2) = "CM2"
                        End If
                        'Del by Zxjike 2013/10/08
                        'If strKataValues(intRI).Contains("6") Or strKataValues(intRI).Contains("CM1") Then
                        '    strSaveType(0) = "6"
                        '    strSaveType(1) = "CM1"
                        'Else
                        '    strSaveType(0) = "8"
                        '    strSaveType(1) = "CM2"
                        'End If
                    End If
                Next
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
        fncGetConectBlockCnt = intConBlockCnt
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncCheckElType
    '*【処理】
    '*   電磁弁形式チェック
    '********************************************************************************************
    Public Shared Function fncCheckElType(objKtbnStrc As KHKtbnStrc, ByVal intNo As Integer, ByVal strElType As String, _
                                     ByRef strCoordinates As String, Optional intRno As Integer = 0) As Boolean
        Dim bolReturn As Boolean = False
        Try
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            '該当列の電磁弁形式をチェック
            For intRI As Integer = Siyou_05.ElType1 - 1 To Siyou_05.ElType6 - 1
                If objKtbnStrc.strcSelection.strPositionInfo(intRI)(intNo) = "1" Then
                    strCoordinates = (intRno.ToString & strComma & CStr(intNo + 1))
                    'Add by Zxjike 2013/10/08 ↓
                    Select Case strElType
                        Case "6"
                            If strKataValues(intRI).StartsWith("PV5-6") Or _
                                strKataValues(intRI).StartsWith("PV5G-6") Or _
                                strKataValues(intRI).StartsWith("CM1") Then
                                bolReturn = True
                                Exit For
                            End If
                        Case "8"
                            If strKataValues(intRI).StartsWith("PV5-8") Or _
                                strKataValues(intRI).StartsWith("PV5G-8") Or _
                                strKataValues(intRI).StartsWith("CM2") Then
                                bolReturn = True
                                Exit For
                            End If
                    End Select
                    'Add by Zxjike 2013/10/08 ↑
                    'Del by Zxjike 2013/10/08
                    'If Left(strKataValues(intRI), 3) = "PV5" And _
                    '   strKataValues(intRI).Contains(strElType) Then
                    '    bolReturn = True
                    '    Exit For
                    'End If
                End If
            Next
            fncCheckElType = bolReturn
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    'Add by Zxjike 2013/10/08 ↓
    Public Shared Function fncCheckMP(objKtbnStrc As KHKtbnStrc, ByVal intNo As Integer, _
                                     ByRef strCoordinates As String) As Boolean
        fncCheckMP = False
        Try
            '該当列の電磁弁形式をチェック
            For intRI As Integer = Siyou_05.ElType1 - 1 To Siyou_05.ElType6 - 1
                If objKtbnStrc.strcSelection.strPositionInfo(intRI)(intNo) = "1" Then
                    strCoordinates = ("0" & strComma & CStr(intNo + 1))
                    If objKtbnStrc.strcSelection.strOptionKataban(intRI).StartsWith("CM1") Or _
                        objKtbnStrc.strcSelection.strOptionKataban(intRI).StartsWith("CM2") Then
                        Exit Function
                    End If
                End If
            Next
            fncCheckMP = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function
End Class
