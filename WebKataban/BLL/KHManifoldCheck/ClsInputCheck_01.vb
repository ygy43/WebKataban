Imports Microsoft.VisualBasic
Imports WebKataban.ClsCommon
Imports WebKataban.CdCst

Public Class ClsInputCheck_01

    Public Shared intColCnt As Integer = 40         'RM1803032_マニホールド連数拡張
    Public Shared intPosRowCnt As Integer = 20

    Public Shared Function fncInputChk(objKtbnStrc As KHKtbnStrc, HT_Option As Hashtable, dblStdNum As Double, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        fncInputChk = False
        Try
            '入力チェック
            If Not fncInpCheck1(objKtbnStrc, HT_Option, dblStdNum, strMsg, strMsgCd) Then
                '画面作成
                Exit Function
            End If

            '入力チェック2
            If Not fncInpCheck2(objKtbnStrc, HT_Option, strMsg, strMsgCd) Then
                '画面作成
                Exit Function
            End If

            '入力チェック2
            If Not fncInpCheck3(objKtbnStrc, strMsg, strMsgCd) Then
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
    '*   入力チェック
    '*【引数】
    '*   strKataValues  : 形番の選択値配列          strUseValues    : 使用数の入力値配列
    '*   arySelectInf   : 設置位置の選択値配列      dblStdNum       : 取付レール長さ（計算値）
    '********************************************************************************************
    Public Shared Function fncInpCheck1(objKtbnStrc As KHKtbnStrc, HT_Option As Hashtable, dblStdNum As Double, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean

        Dim intPosCnt As Integer
        Dim intLoop As Integer
        Dim sbCoordinates As New System.Text.StringBuilder
        Dim strCoordinates As String = String.Empty
        Dim strSaveCoord As String = Nothing
        Dim bolFlag As Boolean
        Dim bolProc As Boolean

        Dim strSwitchPos As String = HT_Option("strSwitchPos")
        Dim strConCaliber As String = HT_Option("strConCaliber")
        Dim strMaxSeq As String = HT_Option("strMaxSeq")
        Dim strOptionT As String = HT_Option("strOptionT")
        Dim strOptionD As String = HT_Option("strOptionD")
        Dim strOptionP7 As String = HT_Option("strOptionP7")

        fncInpCheck1 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            '設置位置が選択されている行の形番が未選択の場合、エラー(添付品)
            For intRI As Integer = 1 To Siyou_01.Inspect4
                '取付レールは除く
                If intRI <> Siyou_01.Rail Then
                    If Int(strUseValues(intRI - 1)) > 0 And _
                       Len(Trim(strKataValues(intRI - 1))) = 0 Then
                        strMsg = intRI & ",0"
                        strMsgCd = "W1400"
                        Exit Function
                    End If
                End If
            Next

            For intCI As Integer = 0 To intColCnt - 1
                '一つの列で３個以上選択されていたらエラー
                intPosCnt = 0
                For intRI As Integer = 0 To intPosRowCnt - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        intPosCnt = intPosCnt + 1
                        sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                    End If
                Next
                If intPosCnt > 2 Then
                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                    strMsg = strCoordinates
                    strMsgCd = "W1390"
                    Exit Function
                End If
                '選択セルの情報を保持しておく
                If Len(sbCoordinates.ToString) > 0 Then
                    strSaveCoord = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                End If
                sbCoordinates = New System.Text.StringBuilder

                '個別配線が選択されている場合
                If arySelectInf(Siyou_01.Wiring - 1)(intCI) = "1" Then
                    'ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ以外が選択されていたらエラー
                    bolFlag = False
                    For intRI As Integer = Siyou_01.Elect1 - 1 To Siyou_01.Elect2 - 1
                        If arySelectInf(intRI)(intCI) = "1" Then
                            bolFlag = True
                            sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                        End If
                    Next
                    For intRI As Integer = Siyou_01.Exhaust1 - 1 To Siyou_01.EndR - 1
                        If arySelectInf(intRI)(intCI) = "1" Then
                            bolFlag = True
                            sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                        End If
                    Next
                    If bolFlag Then
                        strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                        strMsg = strCoordinates
                        strMsgCd = "W1350"
                        Exit Function
                    End If
                    sbCoordinates = New System.Text.StringBuilder

                    'ﾊﾞﾙﾌﾞﾌﾞﾛｯｸが選択されていなかったらエラー
                    For intRI As Integer = Siyou_01.Valve1 - 1 To Siyou_01.Valve7 - 1
                        If arySelectInf(intRI)(intCI) = "1" Then
                            bolFlag = True
                        End If
                    Next
                    If Not bolFlag Then
                        If strOptionT Is Nothing Then
                            strCoordinates = CStr(Siyou_01.Wiring) & strComma & CStr(intCI + 1)
                            strMsg = strCoordinates
                            strMsgCd = "W1360"
                            Exit Function
                        End If
                    End If

                End If

                'ﾊﾞﾙﾌﾞﾌﾞﾛｯｸが選択されている場合
                bolProc = False
                intLoop = Siyou_01.Valve1 - 1
                Do While intLoop < Siyou_01.Valve7
                    If arySelectInf(intLoop)(intCI) = "1" Then
                        bolProc = True
                        Exit Do
                    End If
                    intLoop = intLoop + 1
                Loop
                If bolProc Then
                    bolFlag = False
                    '個別配線が選択されていなかったらエラー
                    If arySelectInf(Siyou_01.Wiring - 1)(intCI) = "0" And _
                       strOptionT Is Nothing Then
                        strCoordinates = CStr(intLoop + 1) & strComma & CStr(intCI + 1)
                        strMsg = strCoordinates
                        strMsgCd = "W1370"
                        Exit Function
                    End If

                    '個別配線以外が選択されていたらエラー
                    For intRI2 As Integer = 0 To Siyou_01.Elect2 - 1
                        If arySelectInf(intRI2)(intCI) = "1" Then
                            bolFlag = True
                            sbCoordinates.Append(CStr(intRI2 + 1) & strComma & CStr(intCI + 1) & strPipe)
                        End If
                    Next
                    For intRI2 As Integer = Siyou_01.Valve1 - 1 To intPosRowCnt - 1
                        If arySelectInf(intRI2)(intCI) = "1" And intLoop <> intRI2 Then
                            bolFlag = True
                            sbCoordinates.Append(CStr(intRI2 + 1) & strComma & CStr(intCI + 1) & strPipe)
                        End If
                    Next
                    If bolFlag Then
                        strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                        strMsg = strCoordinates
                        strMsgCd = "W1380"
                        Exit Function
                    End If
                    sbCoordinates = New System.Text.StringBuilder

                ElseIf arySelectInf(Siyou_01.Wiring - 1)(intCI) = "0" Then
                    'ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ、個別配線が選択されていない状態で２個以上選択されていたらエラー
                    If intPosCnt > 1 Then
                        strMsg = strCoordinates
                        strMsgCd = "W1390"
                        Exit Function
                    End If
                End If
            Next

            'ﾌﾞﾗﾝｸﾌﾟﾗｸﾞ&ｻｲﾚﾝｻ、検査成績所の使用数チェック
            If Not SiyouBLL.fncOtherKataCheck(objKtbnStrc, Siyou_01.Plug1, Siyou_01.Inspect4, _
                                     Siyou_01.Rail, strMsgCd, 99) Then
                Exit Function
            End If

            'ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ形番リスト重複チェック
            For intRI As Integer = Siyou_01.Valve1 - 1 To Siyou_01.Valve7 - 1
                For intRI2 As Integer = intRI + 1 To Siyou_01.Valve7 - 1
                    If Len(strKataValues(intRI)) = 0 Then
                    ElseIf strKataValues(intRI) = strKataValues(intRI2) Then
                        strMsgCd = "W1330"
                        Exit Function
                    End If
                Next
            Next

            '給排気ﾌﾞﾛｯｸ形番リスト重複チェック
            For intRI As Integer = Siyou_01.Exhaust1 - 1 To Siyou_01.Exhaust4 - 1
                For intRI2 As Integer = intRI + 1 To Siyou_01.Exhaust4 - 1
                    If Len(strKataValues(intRI)) = 0 Then
                    ElseIf strKataValues(intRI) = strKataValues(intRI2) Then
                        strMsgCd = "W1330"
                        Exit Function
                    End If
                Next
            Next

            'ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ形番リスト重複チェック
            If Len(strKataValues(Siyou_01.Regulat1 - 1)) = 0 Then
            ElseIf strKataValues(Siyou_01.Regulat1 - 1) = strKataValues(Siyou_01.Regulat2 - 1) Then
                strMsgCd = "W1330"
                Exit Function
            End If

            '取付レール長さチェック
            If strKataValues(Siyou_01.Rail - 1).ToString.Length <= 0 Then strKataValues(Siyou_01.Rail - 1) = 0
            If Not SiyouBLL.fncRailchk(strKataValues(Siyou_01.Rail - 1), CDbl(strUseValues(Siyou_01.Rail - 1)), dblStdNum, strMsgCd) Then
                strMsg = Siyou_01.Rail & ",0"
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
    '*   入力チェック
    '*【引数】
    '*   strKataValues  : 形番の選択値配列          strUseValues    : 使用数の入力値配列
    '*   arySelectInf   : 設置位置の選択値配列
    '********************************************************************************************
    Public Shared Function fncInpCheck2(objKtbnStrc As KHKtbnStrc, HT_Option As Hashtable,
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        Dim intCnt As Integer = 0
        Dim intUse1 As Integer = 0
        Dim intUse2 As Integer = 0
        Dim intLEdge As Integer = 0
        Dim intREdge As Integer = 0
        Dim intLoop As Integer = 0
        Dim intLoop2 As Integer = 0
        Dim intMaxSolD As Integer
        Dim intMaxSolLR As Integer
        Dim intElectSeq As Integer
        Dim intErrRow As Integer
        Dim bolProc As Boolean
        Dim bolMixSwtch(17) As Boolean
        Dim bolFlag As Boolean    'ｼﾝｸﾞﾙｿﾚﾉｲﾄﾞ電磁弁指定チェック
        Dim bolFlag2 As Boolean   '４ﾎﾟｰﾄ電磁弁指定チェック
        Dim bolFlag3 As Boolean
        Dim bolFlag4 As Boolean   '10mmタイプ指定チェック
        Dim bolFlag5 As Boolean   '3ﾎﾟｰﾄ電磁弁指定チェック
        Dim bolFlag6 As Boolean   '7mmタイプ指定チェック
        Dim bolRegConb As Boolean
        Dim strKataban As String
        Dim sbCoordinates As New System.Text.StringBuilder
        Dim strCoordinates As String = ""
        Dim strOption As String = ""
        Dim EOR As Integer  'MN3Q0シリーズ専用
        Dim intMaxSolLRL(1) As Integer  'ソレノイド点数カウント
        Dim intMaxSolLRR(1) As Integer  'ソレノイド点数カウント
        Dim LeftCnt As Integer
        Dim RightCnt As Integer

        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban
        Dim strSwitchPos As String = HT_Option("strSwitchPos")
        Dim strConCaliber As String = HT_Option("strConCaliber")
        Dim strMaxSeq As String = HT_Option("strMaxSeq")
        Dim strOptionT As String = HT_Option("strOptionT")
        Dim strOptionD As String = HT_Option("strOptionD")
        Dim strOptionP7 As String = HT_Option("strOptionP7")

        fncInpCheck2 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            For intRI As Integer = 1 To strUseValues.Count
                If Int(strUseValues(intRI - 1)) > 0 Then
                    '************ 使用OP数カウント(7.1) ************************************
                    Select Case intRI
                        Case Siyou_01.Elect1, Siyou_01.Elect2
                            intCnt = intCnt + 1
                        Case Siyou_01.Valve1, Siyou_01.Valve2, Siyou_01.Valve3, Siyou_01.Valve4, _
                             Siyou_01.Valve5, Siyou_01.Valve6, Siyou_01.Valve7
                            intUse1 = 0
                            intUse2 = 0
                            For intCI As Integer = 1 To intColCnt
                                If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                                    If arySelectInf(Siyou_01.Wiring - 1)(intCI - 1) = "1" Then
                                        intUse1 = intUse1 + 1
                                    Else
                                        intUse2 = intUse2 + 1
                                    End If
                                End If
                            Next
                            If intUse1 > 0 Then
                                intCnt = intCnt + 1
                            End If
                            If intUse2 > 0 Then
                                intCnt = intCnt + 1
                            End If
                        Case Siyou_01.EndL, Siyou_01.EndR
                            intCnt = intCnt + 1
                            '************ ｴﾝﾄﾞﾌﾞﾛｯｸ複数選択チェック(7.5, 7.6) *************
                            If Int(strUseValues(intRI - 1)) > 1 Then
                                For intCI As Integer = 1 To intColCnt
                                    If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                                        sbCoordinates.Append(intRI & strComma & intCI & strPipe)
                                    End If
                                Next
                                strCoordinates = Left(sbCoordinates.ToString, sbCoordinates.ToString.Length - 1)
                                strMsg = strCoordinates
                                strMsgCd = "W1100"
                                Exit Function
                            End If
                        Case Siyou_01.Exhaust1, Siyou_01.Exhaust2, Siyou_01.Exhaust3, Siyou_01.Exhaust4, _
                             Siyou_01.Regulat1, Siyou_01.Regulat2, Siyou_01.Dummy1, Siyou_01.Dummy2
                            intCnt = intCnt + 1
                    End Select
                End If
            Next
            'If intCnt > 20 Then
            '    strMsgCd = "W1010"
            '    Exit Function
            'End If

            '************ 電装ﾌﾞﾛｯｸ必須選択チェック(7.3) ************************************
            If strOptionT = "TX" Then
                If Int(strUseValues(Siyou_01.Elect1 - 1)) = 0 Then
                    sbCoordinates.Append(Siyou_01.Elect1 & strComma & "0")
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1040"
                    Exit Function
                ElseIf Int(strUseValues(Siyou_01.Elect2 - 1)) = 0 Then
                    sbCoordinates.Append(Siyou_01.Elect2 & strComma & "0")
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1040"
                    Exit Function
                ElseIf Int(strUseValues(Siyou_01.Elect1 - 1)) + _
                       Int(strUseValues(Siyou_01.Elect2 - 1)) > 2 Then
                    For intRI As Integer = Siyou_01.Elect1 - 1 To Siyou_01.Elect2 - 1
                        For intCI As Integer = 0 To intColCnt - 1
                            If arySelectInf(intRI)(intCI) = "1" Then
                                sbCoordinates.Append(intRI + 1 & strComma & intCI + 1 & strPipe)
                            End If
                        Next
                    Next
                    strCoordinates = Left(sbCoordinates.ToString, sbCoordinates.ToString.Length - 1)
                    strMsg = strCoordinates
                    strMsgCd = "W1050"
                    Exit Function
                End If
            Else
                If strOptionT IsNot Nothing Then
                    If strUseValues(Siyou_01.Elect1 - 1) = 0 And _
                       strUseValues(Siyou_01.Elect2 - 1) = 0 Then
                        sbCoordinates.Append(Siyou_01.Elect1 & strComma & "0" & strPipe)
                        sbCoordinates.Append(Siyou_01.Elect2 & strComma & "0")
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W1060"
                        Exit Function
                    End If
                    If Int(strUseValues(Siyou_01.Elect1 - 1)) + _
                       Int(strUseValues(Siyou_01.Elect2 - 1)) > 1 Then

                        For intRI As Integer = Siyou_01.Elect1 - 1 To Siyou_01.Elect2 - 1
                            For intCI As Integer = 0 To intColCnt - 1
                                If arySelectInf(intRI)(intCI) = "1" Then
                                    sbCoordinates.Append(intRI + 1 & strComma & intCI + 1 & strPipe)
                                End If
                            Next
                        Next
                        strCoordinates = Left(sbCoordinates.ToString, sbCoordinates.ToString.Length - 1)
                        strMsg = strCoordinates
                        strMsgCd = "W1070"
                        Exit Function
                    End If
                End If
            End If

            bolFlag = False
            For intCI As Integer = 0 To intColCnt - 1
                '最左端と最右端の列数を取得
                intLoop = 0
                Do While intLoop < intPosRowCnt
                    If arySelectInf(intLoop)(intCI) = "1" Then
                        If intLEdge = 0 Then
                            intLEdge = intCI + 1
                            intREdge = intCI + 1
                        End If
                        '中間に全く選択されていない列が存在する場合、エラー(7.2)
                        If intREdge < intCI Then
                            For intI As Integer = intREdge + 1 To intCI
                                sbCoordinates.Append(intI & strPipe)
                                strOption = strOption & CStr(intI) & strComma
                                Exit For
                            Next
                            strCoordinates = Left(sbCoordinates.ToString, sbCoordinates.ToString.Length - 1)
                            strOption = Left(strOption, strOption.Length - 1)
                            strMsg = strCoordinates
                            strMsgCd = "W1020"
                            Exit Function
                        Else
                            intREdge = intCI + 1
                        End If
                        Exit Do
                    End If
                    intLoop = intLoop + 1
                Loop

                '************ 個別配線のチェック(7.4) **************************************
                Select Case strSeriesKata
                    Case "MN3Q0", "MT3Q0"
                        'MN3Q0,MT3Q0シリーズは個別配線がないためチェックなし
                    Case Else
                        If arySelectInf(Siyou_01.Wiring - 1)(intCI) = "1" Then
                            bolFlag = True
                            '個別配線指定で、バルブブロックが指定されていない場合、エラー
                            intLoop = Siyou_01.Valve1 - 1
                            Do While intLoop <= Siyou_01.Valve7
                                If intLoop = Siyou_01.Valve7 Then
                                    sbCoordinates.Append(Siyou_01.Wiring & strComma & intCI + 1)
                                    strMsg = sbCoordinates.ToString
                                    strMsgCd = "W1080"
                                    Exit Function
                                ElseIf arySelectInf(intLoop)(intCI) = "1" Then
                                    Exit Do
                                End If
                                intLoop = intLoop + 1
                            Loop
                        ElseIf strOptionT Is Nothing And _
                               strOptionD IsNot Nothing Then
                            '個別配線のみで、バルブブロックに個別配線が指定されていない場合、エラー
                            If strSeriesKata <> "MN3EX0" Or strSeriesKata <> "MN4EX0" Then
                                For intRI As Integer = Siyou_01.Valve1 - 1 To Siyou_01.Valve7 - 1

                                    If arySelectInf(intRI)(intCI) = "1" Then
                                        For intI As Integer = Siyou_01.Valve1 - 1 To Siyou_01.Valve7 - 1
                                            sbCoordinates.Append(intI + 1 & strComma & intCI + 1 & strPipe)
                                        Next
                                        strCoordinates = Left(sbCoordinates.ToString, sbCoordinates.ToString.Length - 1)
                                        strMsg = strCoordinates
                                        strMsgCd = "W1090"
                                        Exit Function
                                    End If
                                Next
                            End If
                        End If
                End Select

                Select Case strSeriesKata
                    Case "MN3Q0"
                        If strOptionT = "TX" Then
                            EOR = 0
                            Call MN3Q0_Error(objKtbnStrc, EOR, LeftCnt, RightCnt)

                            If EOR = 1 Then
                                strMsg = strCoordinates
                                strMsgCd = "W8890"
                                Exit Function
                            End If
                        End If
                        '電装ブロックミックス(TX)T**R配線最終端指示の複数指定チェック
                        If strUseValues(Siyou_01.Wiring - 1) > 1 Then
                            sbCoordinates.Append(CStr(Siyou_01.Wiring) & strComma & "0")
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W8880"
                            Exit Function
                        End If
                End Select

                Select Case strSeriesKata
                    Case "MT3Q0"
                        '給排気ブロック１の複数指定チェック
                        If strUseValues(Siyou_01.Exhaust1 - 1) > 1 Then
                            sbCoordinates.Append(CStr(Siyou_01.Exhaust1) & strComma & "0")
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W8880"
                            Exit Function
                        End If
                        '給排気ブロック２の複数指定チェック
                        If strUseValues(Siyou_01.Exhaust2 - 1) > 1 Then
                            sbCoordinates.Append(CStr(Siyou_01.Exhaust2) & strComma & "0")
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W8880"
                            Exit Function
                        End If
                        '給排気ブロック３の複数指定チェック
                        If strUseValues(Siyou_01.Exhaust3 - 1) > 1 Then
                            sbCoordinates.Append(CStr(Siyou_01.Exhaust3) & strComma & "0")
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W8880"
                            Exit Function
                        End If
                End Select

                ' MN3EX0，MN4EX0においてバルブブロックに７ｍｍの箇所に個別配線を指定した場合エラー
                Select Case strSeriesKata
                    Case "MN3EX0", "MN4EX0"
                        If Len(strOptionD) <> 0 Then
                            For intRI As Integer = Siyou_01.Valve1 - 1 To Siyou_01.Valve7 - 1
                                If arySelectInf(intRI)(intCI) = "1" Then
                                    Select Case intRI + 1
                                        Case Siyou_01.Valve1, Siyou_01.Valve2, Siyou_01.Valve3, Siyou_01.Valve4, _
                                         Siyou_01.Valve5, Siyou_01.Valve6, Siyou_01.Valve7
                                            strKataban = Trim(strKataValues(intRI))
                                            Select Case Mid(strKataban, 1, 5)
                                                Case "N3E00", "N4E00"
                                                    If arySelectInf(Siyou_01.Wiring - 1)(intCI) = "1" Then
                                                        If objKtbnStrc.strcSelection.strOpSymbol(6) = "E" Or _
                                                            objKtbnStrc.strcSelection.strOpSymbol(10) = "ST" Then
                                                            strMsgCd = "W8760"
                                                            '"7mmタイプのバルブに個別配線は指定できません。"
                                                            sbCoordinates.Append(Siyou_01.Wiring & strComma & intCI)
                                                            strMsg = sbCoordinates.ToString
                                                            Exit Function
                                                        End If
                                                    End If
                                            End Select
                                    End Select
                                End If
                            Next
                        End If
                End Select
            Next
            '全体で一つも選択されていない場合、エラー(7.2)
            If intLEdge = 0 Then
                strMsgCd = "W1030"
                Exit Function
            End If

            Select Case strSeriesKata
                Case "MN3Q0", "MT3Q0"
                    'MN3Q0,MT3Q0シリーズは個別配線がないためチェックなし
                Case Else
                    '個別配線MIX指定で、個別配送が指定されていない場合、エラー
                    If strOptionT IsNot Nothing And _
                       strOptionD IsNot Nothing And _
                       Not bolFlag Then
                        strCoordinates = Siyou_01.Wiring & strComma & "0"
                        strMsg = strCoordinates
                        strMsgCd = "W1090"
                        Exit Function
                    End If
            End Select

            For intRI As Integer = 0 To intPosRowCnt - 1
                '************ マニホールド最左部セルのチェック(7.7) **************************
                If arySelectInf(intRI)(0) = "1" Then
                    If intRI = Siyou_01.Elect1 - 1 Or intRI = Siyou_01.Elect2 - 1 Then
                        If strKataValues(intRI).Contains("R") Or strKataValues(intRI).Contains("TM") Then
                            strCoordinates = "0" & strComma & "1"
                            strMsg = strCoordinates
                            strMsgCd = "W1110"
                            Exit Function
                        End If
                    ElseIf intRI <> Siyou_01.EndL - 1 Then
                        strCoordinates = "0" & strComma & "1"
                        strMsg = strCoordinates
                        strMsgCd = "W1110"
                        Exit Function
                    End If
                End If

                '************ マニホールド最右部セルのチェック(7.8) ***************************
                If arySelectInf(intRI)(intREdge - 1) = "1" Then
                    If intRI = Siyou_01.Elect1 - 1 Or intRI = Siyou_01.Elect2 - 1 Then
                        If Not strKataValues(intRI).Contains("R") Or strKataValues(intRI).Contains("TM") Then
                            strCoordinates = intRI + 1 & strComma & intREdge
                            strMsg = strCoordinates
                            strMsgCd = "W1120"
                            Exit Function
                        End If
                    ElseIf intRI <> Siyou_01.EndR - 1 Then
                        strCoordinates = intRI + 1 & strComma & intREdge
                        strMsg = strCoordinates
                        strMsgCd = "W1120"
                        Exit Function
                    End If

                End If
            Next

            '************ ソレノイド点数チェック(7.9) 左 → 右 ******************************************
            If strOptionT Is Nothing Then
                intMaxSolD = 48
            Else
                'If objKtbnStrc.strcSelection.strOpSymbol(6) = "T30" Then
                '    intMaxSolD = 24
                'Else
                '    intMaxSolD = 16
                'End If

                intMaxSolD = KHKataban.fncGetMaxSol(objKtbnStrc.strcSelection.strOpSymbol, 19)
            End If

            For intCI As Integer = 0 To intColCnt - 1
                intMaxSolLR = 0
                For intRI As Integer = Siyou_01.Elect1 - 1 To Siyou_01.Elect2 - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        If Not strKataValues(intRI).Contains("R") Then

                            If strOptionT IsNot Nothing Then
                                '確認処理(Ｔ＊)
                                intCnt = fncSolenoidCnt1(objKtbnStrc, intCI + 1, intColCnt - 1, 1)
                                If intCnt = 0 Then
                                    If strOptionT = "TX" Then
                                        strMsgCd = "W1130"
                                        Exit Function
                                    Else
                                        strMsgCd = "W1140"
                                        Exit Function
                                    End If
                                End If
                            End If

                            '確認処理(Ｄ＊)
                            intCnt = fncSolenoidCnt2(objKtbnStrc, intCI + 1, intColCnt - 1, 1, intMaxSolD)
                            If intCnt > intMaxSolD Then
                                strMsgCd = "W1150"
                                Exit Function
                            End If

                            If strSeriesKata = "MN3Q0" Or strSeriesKata = "MT3Q0" Then
                                If strOptionT = "TX" Then
                                    For i As Integer = 0 To 1
                                        intMaxSolLRL(i) = KHKataban.fncGetMaxSol_01(strKataValues(i))
                                    Next
                                    'ソレノイドＭＡＸ値(電装ブロックＬ)取得
                                    intMaxSolLR = intMaxSolLRL(0) + intMaxSolLRL(1)

                                    LeftCnt = LeftCnt * 2
                                    RightCnt = RightCnt * 2
                                    If LeftCnt > intMaxSolLRL(0) Or RightCnt > intMaxSolLRL(1) Then
                                        strMsgCd = "W8890"
                                        Exit Function
                                    End If

                                Else
                                    'ソレノイドＭＡＸ値(電装ブロックＬ)取得
                                    intMaxSolLR = KHKataban.fncGetMaxSol_01(strKataValues(intRI))
                                End If
                            Else
                                'ソレノイドＭＡＸ値(電装ブロックＬ)取得
                                intMaxSolLR = KHKataban.fncGetMaxSol_01(strKataValues(intRI))
                            End If

                            '確認処理(電装ブロックＬ)
                            intCnt = fncSolenoidCnt3(objKtbnStrc, intCI + 1, intColCnt - 1, 1, intMaxSolLR)
                            If intCnt > intMaxSolLR Then
                                strMsgCd = "W1150"
                                Exit Function
                            End If
                        End If
                        If strKataValues(intRI).Contains("TM") Then
                            '********* 電装ﾌﾞﾛｯｸ ＆ 給排気ﾌﾞﾛｯｸ ＆ ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ組み合わせチェック(7.12,7.13) **************
                            subConbCheck1(objKtbnStrc, bolFlag, bolFlag2, bolFlag3, intCI - 1)
                            If bolFlag Then
                                If bolFlag2 Or Not bolFlag3 Then
                                    strMsgCd = "W1240"
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next

                bolProc = False
                intLoop = Siyou_01.Exhaust1 - 1
                Do While intLoop < Siyou_01.Exhaust4
                    If arySelectInf(intLoop)(intCI) = "1" Then
                        If strKataValues(intLoop).Contains("-S") Then
                            bolProc = True
                            Exit Do
                        End If
                    End If
                    intLoop = intLoop + 1
                Loop
                If bolProc Then

                    '************ 給排気ﾌﾞﾛｯｸ ＆ 電装ﾌﾞﾛｯｸ組み合わせチェック(7.14) ***************************
                    If Not fncCheckProc14(objKtbnStrc, intCI - 1, intErrRow) Then
                        sbCoordinates.Append("0" & strComma & intCI + 1 & strPipe)
                        sbCoordinates.Append(intErrRow & strComma & intCI)
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W1250"
                        Exit Function
                    End If

                    '************ 給排気ﾌﾞﾛｯｸ ＆ ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ組み合わせチェック(7.15) ***************************
                    If Not fncConbCheck2(objKtbnStrc, intCI - 1, 0, -1, strCoordinates) Then
                        sbCoordinates.Append("0" & strComma & intCI + 1 & strPipe)
                        sbCoordinates.Append(strCoordinates)
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W1250"
                        Exit Function
                    End If
                    If Not fncConbCheck2(objKtbnStrc, intCI + 1, intColCnt - 1, 1, strCoordinates) Then
                        sbCoordinates.Append("0" & strComma & intCI + 1 & strPipe)
                        sbCoordinates.Append(strCoordinates)
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W1250"
                        Exit Function
                    End If
                End If

                bolProc = False
                intLoop = Siyou_01.Exhaust1 - 1
                Do While intLoop < Siyou_01.Exhaust4
                    If arySelectInf(intLoop)(intCI) = "1" Then
                        bolProc = True
                        Exit Do
                    End If
                    intLoop = intLoop + 1
                Loop
                If bolProc Then

                    '************ 給排気ﾌﾞﾛｯｸ ＆ ｴﾝﾄﾞﾌﾞﾛｯｸ組み合わせチェック(7.20) ***************************
                    If strKataValues(intLoop).Contains("-SA") Then
                        subConbCheck4(objKtbnStrc, bolFlag, bolFlag2, bolFlag3, intCI + 1, intColCnt - 1, 1, strCoordinates)
                        If bolFlag And (Not bolFlag2) And (Not bolFlag3) Then
                            strMsg = strCoordinates
                            strMsgCd = "W1260"
                            Exit Function
                        End If
                    End If
                End If

                '************ ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ ＆ 給排気ﾌﾞﾛｯｸ組み合わせチェック(7.33) ***************************
                bolProc = False
                intLoop = Siyou_01.Regulat1 - 1
                Do While intLoop < Siyou_01.Regulat2
                    If arySelectInf(intLoop)(intCI) = "1" And _
                       fncContaints(strKataValues(intLoop), "RA-LR,RA-FR,RB-LR,RB-FR") Then
                        bolProc = True
                        Exit Do
                    End If
                    intLoop = intLoop + 1
                Loop
                If bolProc Then
                    If Not fncCheckProc33(objKtbnStrc, intCI + 1, strCoordinates) Then
                        strMsg = strCoordinates
                        strMsgCd = "W1290"
                        Exit Function
                    End If

                End If

                '************ ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ ＆ ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ ＆ 給排気ﾌﾞﾛｯｸ組み合わせチェック(7.34) ***************************
                bolProc = False
                intLoop = Siyou_01.Regulat1 - 1
                Do While intLoop < Siyou_01.Regulat2
                    If arySelectInf(intLoop)(intCI) = "1" And _
                       fncContaints(strKataValues(intLoop), "-LR,-FR") Then
                        bolProc = True
                        Exit Do
                    End If
                    intLoop = intLoop + 1
                Loop
                If bolProc Then
                    If Not fncCheckProc34(objKtbnStrc, intCI - 1, sbCoordinates) Then
                        strCoordinates = sbCoordinates.ToString & CStr(intLoop + 1) & strComma & CStr(intCI + 1)
                        strMsg = strCoordinates
                        strMsgCd = "W1290"
                        Exit Function
                    End If
                    sbCoordinates = New System.Text.StringBuilder
                End If

                '************ ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ ＆ 給排気ﾌﾞﾛｯｸ ＆ ｴﾝﾄﾞﾌﾞﾛｯｸ ＆ 電装ﾌﾞﾛｯｸ組み合わせチェック(7.35) ***************************
                intLoop = Siyou_01.Regulat1 - 1
                Do While intLoop < Siyou_01.Regulat2
                    If arySelectInf(intLoop)(intCI) = "1" Then

                        If fncContaints(strKataValues(intLoop), "-FL,-LR,-RL") Then
                            For intRI As Integer = Siyou_01.Exhaust1 - 1 To Siyou_01.Exhaust4 - 1
                                If arySelectInf(intRI)(intCI - 1) = "1" And _
                                   strKataValues(intRI).Contains("-S") Then
                                    sbCoordinates.Append(CStr(intLoop + 1) & strComma & CStr(intCI + 1) & strPipe)
                                    sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI))
                                    strMsg = sbCoordinates.ToString
                                    strMsgCd = "W1260"
                                    Exit Function
                                End If
                            Next

                            For intRI As Integer = Siyou_01.Elect1 - 1 To Siyou_01.Elect2 - 1
                                If arySelectInf(intRI)(intCI - 1) = "1" And _
                                   strKataValues(intRI).Contains("R") Then
                                    sbCoordinates.Append(CStr(intLoop + 1) & strComma & CStr(intCI + 1) & strPipe)
                                    sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI))
                                    strMsg = sbCoordinates.ToString
                                    strMsgCd = "W1260"
                                    Exit Function
                                End If
                            Next
                            If arySelectInf(Siyou_01.EndL - 1)(intCI - 1) = "1" And _
                               strKataValues(Siyou_01.EndL - 1).Contains("-EL") Then
                                sbCoordinates.Append(CStr(intLoop + 1) & strComma & CStr(intCI + 1) & strPipe)
                                sbCoordinates.Append(CStr(Siyou_01.EndL) & strComma & CStr(intCI))
                                strMsg = sbCoordinates.ToString
                                strMsgCd = "W1260"
                                Exit Function
                            End If
                        End If

                        If fncContaints(strKataValues(intLoop), "-FR,-LR,-RL") Then
                            For intRI As Integer = Siyou_01.Elect1 - 1 To Siyou_01.Elect2 - 1
                                If arySelectInf(intRI)(intCI + 1) = "1" And _
                                   strKataValues(intRI).Contains("R") Then
                                    sbCoordinates.Append(CStr(intLoop + 1) & strComma & CStr(intCI + 1) & strPipe)
                                    sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 2))
                                    strMsg = sbCoordinates.ToString
                                    strMsgCd = "W1260"
                                    Exit Function
                                End If
                            Next
                            If arySelectInf(Siyou_01.EndR - 1)(intCI + 1) = "1" And _
                               strKataValues(Siyou_01.EndR - 1).Contains("-ER") Then
                                sbCoordinates.Append(CStr(intLoop + 1) & strComma & CStr(intCI + 1) & strPipe)
                                sbCoordinates.Append(CStr(Siyou_01.EndR) & strComma & CStr(intCI + 2))
                                strMsg = sbCoordinates.ToString
                                strMsgCd = "W1260"
                                Exit Function
                            End If

                            '************ ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ組み合わせチェック(7.36) ***************************
                            For intRI As Integer = Siyou_01.Regulat1 - 1 To Siyou_01.Regulat2 - 1
                                If arySelectInf(intRI)(intCI + 1) = "1" And _
                                   fncContaints(strKataValues(intRI), "-FR,-FL,-LR,-RL") Then
                                    sbCoordinates.Append(CStr(intLoop + 1) & strComma & CStr(intCI + 1) & strPipe)
                                    sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 2))
                                    strMsg = sbCoordinates.ToString
                                    strMsgCd = "W1260"
                                    Exit Function
                                End If

                            Next
                        End If

                        If strKataValues(intLoop).Contains("-FL") Then
                            For intRI As Integer = Siyou_01.Regulat1 - 1 To Siyou_01.Regulat2 - 1
                                If arySelectInf(intRI)(intCI + 1) = "1" And _
                                   fncContaints(strKataValues(intRI), "-FL,-LR,-RL") Then
                                    sbCoordinates.Append(CStr(intLoop + 1) & strComma & CStr(intCI + 1) & strPipe)
                                    sbCoordinates.Append(CStr(Siyou_01.EndL) & strComma & CStr(intCI))
                                    strMsg = sbCoordinates.ToString
                                    strMsgCd = "W1260"
                                    Exit Function
                                End If
                            Next
                        End If
                    End If
                    intLoop = intLoop + 1
                Loop
            Next

            '************ 給排気ﾌﾞﾛｯｸ単独チェック(7.16) ***************************
            bolProc = False
            intLoop = Siyou_01.Exhaust1 - 1
            Do While intLoop < Siyou_01.Exhaust4
                If Int(strUseValues(intLoop)) > 0 Then

                    bolProc = True
                    Exit Do
                End If
                intLoop = intLoop + 1
            Loop
            If bolProc Then
                intLoop2 = intLoop
                Do While intLoop2 < Siyou_01.Exhaust4
                    If fncContaints(strKataValues(intLoop2), "-Q-,-QK-,-QKX-,-QKZ-,-QX-") Then
                        Exit Do
                    End If
                    intLoop2 = intLoop2 + 1
                Loop
            Else
                strMsgCd = "W1260"
                Exit Function
            End If
            If intLoop2 = Siyou_01.Exhaust4 Then
                For intI As Integer = Siyou_01.Exhaust1 To Siyou_01.Exhaust4
                    sbCoordinates.Append(intI & strComma & "0" & strPipe)
                Next
                strCoordinates = sbCoordinates.ToString.Substring(0, sbCoordinates.ToString.Length - 1)
                strMsg = strCoordinates
                strMsgCd = "W1260"
                Exit Function
            End If

            '************ 給排気ﾌﾞﾛｯｸ ＆ ｴﾝﾄﾞﾌﾞﾛｯｸ組み合わせチェック(7.21) ***************************
            For intCI As Integer = 0 To intColCnt - 1
                If ((arySelectInf(Siyou_01.Elect1 - 1)(intCI) = "1" And strKataValues(Siyou_01.Elect1 - 1).Contains("R")) Or _
                    (arySelectInf(Siyou_01.Elect2 - 1)(intCI) = "1" And strKataValues(Siyou_01.Elect2 - 1).Contains("R"))) _
                   Or arySelectInf(Siyou_01.EndR - 1)(intCI) = "1" Then
                    subConbCheck4(objKtbnStrc, bolFlag, bolFlag2, bolFlag3, intCI - 1, 0, -1, strCoordinates)
                    If bolFlag And Not bolFlag2 Then
                        strMsg = strCoordinates
                        strMsgCd = "W1260"
                        Exit Function
                    End If
                    Exit For
                End If
            Next

            '************ 給排気ﾌﾞﾛｯｸ ＆ 電装ﾌﾞﾛｯｸ組み合わせチェック２(7.22) ***************************
            For intCI As Integer = 0 To intColCnt - 1
                bolProc = False
                intLoop = Siyou_01.Exhaust1 - 1
                Do While intLoop < Siyou_01.Exhaust4
                    If arySelectInf(intLoop)(intCI) = "1" Then
                        If strKataValues(intLoop).Contains("-SA") And strKataValues(intLoop).Contains("-QZ") Then
                            bolProc = True
                            Exit Do
                        Else
                            Exit For
                        End If
                    End If
                    intLoop = intLoop + 1
                Loop
                If bolProc Then
                    If Not fncCheckProc22(objKtbnStrc, intCI - 1) Then
                        strCoordinates = CStr(intLoop + 1) & strComma & CStr(intCI + 1)
                        strMsg = strCoordinates
                        strMsgCd = "W1260"
                        Exit Function
                    End If
                End If
            Next

            '************ 給排気ﾌﾞﾛｯｸ ＆ ｴﾝﾄﾞﾌﾞﾛｯｸ組み合わせチェック(7.23) ***************************
            For intCI As Integer = 0 To intColCnt - 1
                bolFlag3 = False
                bolProc = False
                intLoop = Siyou_01.Exhaust1 - 1
                Do While intLoop < Siyou_01.Exhaust4
                    If arySelectInf(intLoop)(intCI) = "1" Then
                        If strKataValues(intLoop).Contains("-S") And _
                           Not strKataValues(intLoop).Contains("-SA") And _
                           Not strKataValues(intLoop).Contains("-QZ-") Then

                            bolProc = True
                            Exit Do
                        End If
                    End If
                    intLoop = intLoop + 1
                Loop
                If bolProc Then
                    subCheckProc23(objKtbnStrc, bolFlag, bolFlag2, bolFlag3, intCI + 1, strCoordinates)
                    If bolFlag And Not bolFlag2 And Not bolFlag3 Then
                        sbCoordinates.Append(CStr(intLoop + 1) & strComma & CStr(intCI + 1) & strPipe)
                        sbCoordinates.Append(strCoordinates)
                        strMsg = sbCoordinates.ToString
                        strMsgCd = "W1270"
                        Exit Function
                    End If
                End If
            Next

            If ((strSeriesKata <> "MN3EX0" Or strSeriesKata <> "MN4EX0") And objKtbnStrc.strcSelection.strOpSymbol(4) = "R") Or _
               ((strSeriesKata = "MN3EX0" Or strSeriesKata = "MN4EX0") And objKtbnStrc.strcSelection.strOpSymbol(2) = "R") Then
                '************ ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ選択必須チェック(7.24) ******************************
                bolFlag = False
                For intCI As Integer = 0 To intColCnt - 1
                    If arySelectInf(Siyou_01.Regulat1 - 1)(intCI) = "1" Or _
                       arySelectInf(Siyou_01.Regulat2 - 1)(intCI) = "1" Then
                        bolFlag = True
                    End If
                Next
                If Not bolFlag Then
                    sbCoordinates.Append(Siyou_01.Regulat1 & strComma & "0" & strPipe)
                    sbCoordinates.Append(Siyou_01.Regulat2 & strComma & "0")
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1280"
                    Exit Function
                End If

                bolRegConb = True
                For intCI As Integer = 0 To intColCnt - 1
                    '************ 給排気ﾌﾞﾛｯｸ ＆ ｴﾝﾄﾞﾌﾞﾛｯｸ ＆ 電装ﾌﾞﾛｯｸ組合せチェック(7.17 / 7.18 / 7.19) ******************************
                    bolProc = False
                    If arySelectInf(Siyou_01.EndR - 1)(intCI) = "1" Then
                        bolProc = True
                        intErrRow = Siyou_01.EndR
                    ElseIf arySelectInf(Siyou_01.Elect1 - 1)(intCI) = "1" And _
                           strKataValues(Siyou_01.Elect1 - 1).Contains("R") Then
                        bolProc = True
                        intErrRow = Siyou_01.Elect1
                    ElseIf arySelectInf(Siyou_01.Elect2 - 1)(intCI) = "1" And _
                           strKataValues(Siyou_01.Elect2 - 1).Contains("R") Then
                        bolProc = True
                        intErrRow = Siyou_01.Elect2
                    End If
                    If bolProc Then
                        For intI As Integer = 1 To 2
                            If Not fncConbCheck3(objKtbnStrc, intCI - 1, intI, strCoordinates) Then
                                sbCoordinates.Append(CStr(intErrRow) & strComma & CStr(intCI + 1) & strPipe)
                                sbCoordinates.Append(strCoordinates)
                                strMsg = sbCoordinates.ToString
                                strMsgCd = "W1260"
                                Exit Function
                            End If
                        Next
                    End If

                    '************ ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ ＆ ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ ＆ 給排気ﾌﾞﾛｯｸ組み合わせチェック(7.25) ***************************
                    bolProc = False
                    intLoop = Siyou_01.Regulat1 - 1
                    Do While intLoop < Siyou_01.Regulat2
                        If arySelectInf(intLoop)(intCI) = "1" And _
                           strKataValues(intLoop).Contains("-FL") Then
                            bolProc = True
                            Exit Do
                        End If
                        intLoop = intLoop + 1
                    Loop

                    If bolProc Then
                        subConbCheck5(objKtbnStrc, bolFlag, bolFlag2, bolFlag3, intCI + 1, intColCnt - 1, 1, 1, sbCoordinates)
                        If bolFlag And Not bolFlag2 And Not bolFlag3 Then
                            strCoordinates = CStr(intLoop + 1) & strComma & CStr(intCI + 1)
                            strMsg = strCoordinates
                            strMsgCd = "W1290"
                            Exit Function
                        End If
                        sbCoordinates = New System.Text.StringBuilder
                    Else

                        '************ ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ ＆ ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ ＆ 給排気ﾌﾞﾛｯｸ組み合わせチェック(7.26) ***************************
                        intLoop = Siyou_01.Regulat1 - 1
                        Do While intLoop < Siyou_01.Regulat2
                            If arySelectInf(intLoop)(intCI) = "1" And _
                               fncContaints(strKataValues(intLoop), "-FL,-RL") Then
                                bolProc = True
                                Exit Do
                            End If
                            intLoop = intLoop + 1
                        Loop
                    End If

                    If bolProc Then
                        subConbCheck5(objKtbnStrc, bolFlag, bolFlag2, bolFlag3, intCI - 1, 0, -1, 2, sbCoordinates)
                        If Not bolFlag And Not bolFlag2 Then
                            strCoordinates = CStr(intLoop + 1) & strComma & CStr(intCI + 1)
                            strMsg = strCoordinates
                            strMsgCd = "W1290"
                            Exit Function
                        End If
                        sbCoordinates = New System.Text.StringBuilder
                    End If

                    '************ ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ ＆ ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ ＆ 給排気ﾌﾞﾛｯｸ組み合わせチェック(7.27) ***************************
                    bolProc = False
                    intLoop = Siyou_01.Regulat1 - 1
                    Do While intLoop < Siyou_01.Regulat2
                        If arySelectInf(intLoop)(intCI) = "1" And _
                           strKataValues(intLoop).Contains("-FR") Then

                            bolProc = True
                            Exit Do
                        End If
                        intLoop = intLoop + 1
                    Loop

                    If bolProc Then
                        subConbCheck5(objKtbnStrc, bolFlag, bolFlag2, bolFlag3, intCI - 1, 0, -1, 3, sbCoordinates)
                        If bolFlag And Not bolFlag2 And Not bolFlag3 Then
                            strCoordinates = CStr(intLoop + 1) & strComma & CStr(intCI + 1)
                            strCoordinates = sbCoordinates.ToString & strCoordinates
                            strMsg = strCoordinates
                            strMsgCd = "W1290"
                            Exit Function
                        End If
                        sbCoordinates = New System.Text.StringBuilder
                    Else
                        '************ ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ ＆ ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ ＆ 給排気ﾌﾞﾛｯｸ組み合わせチェック(7.28) ***************************
                        intLoop = Siyou_01.Regulat1 - 1
                        Do While intLoop < Siyou_01.Regulat2
                            If arySelectInf(intLoop)(intCI) = "1" And fncContaints(strKataValues(intLoop), "-FR,-LR") Then
                                bolProc = True
                                Exit Do
                            End If
                            intLoop = intLoop + 1
                        Loop
                    End If

                    If bolProc Then
                        subConbCheck5(objKtbnStrc, bolFlag, bolFlag2, bolFlag3, intCI + 1, intColCnt - 1, 1, 4, sbCoordinates)
                        If Not bolFlag And Not bolFlag2 Then
                            strCoordinates = CStr(intLoop + 1) & strComma & CStr(intCI + 1)
                            strMsg = strCoordinates
                            strMsgCd = "W1290"
                            Exit Function
                        End If
                        sbCoordinates = New System.Text.StringBuilder
                    End If

                    '************ ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ ＆ 給排気ﾌﾞﾛｯｸ組み合わせチェック(7.29) ***************************
                    bolProc = False
                    intLoop = Siyou_01.Regulat1 - 1
                    Do While intLoop < Siyou_01.Regulat2
                        If arySelectInf(intLoop)(intCI) = "1" And strKataValues(intLoop).Contains("-RL") Then
                            bolProc = True
                            Exit Do
                        End If
                        intLoop = intLoop + 1
                    Loop
                    If bolProc Then
                        If Not fncConbCheck6(objKtbnStrc, intCI + 1, intColCnt - 1, 1, 1) Then
                            strCoordinates = CStr(intLoop + 1) & strComma & CStr(intCI + 1)
                            strMsg = strCoordinates
                            strMsgCd = "W1290"
                            Exit Function
                        End If
                    End If

                    '************ ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ ＆ 給排気ﾌﾞﾛｯｸ組み合わせチェック(7.30) ***************************
                    bolProc = False
                    intLoop = Siyou_01.Regulat1 - 1
                    Do While intLoop < Siyou_01.Regulat2
                        If arySelectInf(intLoop)(intCI) = "1" And strKataValues(intLoop).Contains("-LR") Then

                            bolProc = True
                            Exit Do
                        End If
                        intLoop = intLoop + 1
                    Loop
                    If bolProc Then
                        If Not fncConbCheck6(objKtbnStrc, intCI - 1, 0, -1, 1) Then
                            strCoordinates = CStr(intLoop + 1) & strComma & CStr(intCI + 1)
                            strMsg = strCoordinates
                            strMsgCd = "W1290"
                            Exit Function
                        End If
                    End If

                    '************ ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ ＆ 給排気ﾌﾞﾛｯｸ組み合わせチェック(7.31) ***************************
                    If bolRegConb Then
                        bolProc = False
                        intLoop = Siyou_01.Exhaust1 - 1
                        Do While intLoop < Siyou_01.Exhaust4
                            If arySelectInf(intLoop)(intCI) = "1" Then

                                bolProc = True
                                Exit Do
                            End If
                            intLoop = intLoop + 1
                        Loop
                        If bolProc Then
                            If fncConbCheck6(objKtbnStrc, intCI - 1, 0, -1, 2) Then
                                strCoordinates = CStr(intLoop + 1) & strComma & CStr(intCI + 1)
                                strMsg = strCoordinates
                                strMsgCd = "W1290"
                                Exit Function
                            End If
                            bolRegConb = False
                        End If
                    End If

                Next

                '使用数が0以上で形番が未選択の場合(取付レール、チューブ抜具は処理の対象外とする)
                For intRI As Integer = 0 To strUseValues.Count - 1
                    If Int(strUseValues(intRI)) > 0 And intRI <> Siyou_01.Rail - 1 Then
                        If Len(Trim(strKataValues(intRI))) = 0 Then
                            strMsgCd = "W1310"
                            Exit Function
                        End If
                    End If
                Next
            End If

            bolRegConb = True
            '************ ソレノイド点数チェック(7.9) 右 → 左 ******************************************
            For intCI As Integer = intColCnt - 1 To 0 Step -1
                intMaxSolLR = 0
                For intRI As Integer = Siyou_01.Elect1 - 1 To Siyou_01.Elect2 - 1
                    If arySelectInf(intRI)(intCI) = "1" And strKataValues(intRI).Contains("R") Then

                        If strOptionT IsNot Nothing Then
                            '確認処理(Ｔ＊)
                            intCnt = fncSolenoidCnt1(objKtbnStrc, intCI - 1, 0, -1)
                            If intCnt = 0 Then
                                If strOptionT = "TX" Then
                                    strMsgCd = "W1130"
                                    Exit Function
                                Else
                                    strMsgCd = "W1140"
                                    Exit Function
                                End If
                            End If
                        End If

                        '確認処理(Ｄ＊)
                        intCnt = fncSolenoidCnt2(objKtbnStrc, intCI - 1, 0, -1, intMaxSolD)
                        If intCnt > intMaxSolD Then
                            strMsgCd = "W1150"
                            Exit Function
                        End If

                        If strSeriesKata = "MN3Q0" Or strSeriesKata = "MT3Q0" Then
                            If strOptionT = "TX" Then
                                For i As Integer = 0 To 1
                                    intMaxSolLRR(i) = KHKataban.fncGetMaxSol_01(strKataValues(i))
                                Next
                                'ソレノイドＭＡＸ値(電装ブロックＲ)取得
                                intMaxSolLR = intMaxSolLRR(0) + intMaxSolLRR(1)
                            Else
                                'ソレノイドＭＡＸ値(電装ブロックＲ)取得
                                intMaxSolLR = KHKataban.fncGetMaxSol_01(strKataValues(intRI))
                            End If
                        Else
                            'ソレノイドＭＡＸ値(電装ブロックＲ)取得
                            intMaxSolLR = KHKataban.fncGetMaxSol_01(strKataValues(intRI))
                        End If

                        '確認処理(電装ブロックＲ)
                        intCnt = fncSolenoidCnt3(objKtbnStrc, intCI - 1, 0, -1, intMaxSolLR)
                        If intCnt > intMaxSolLR Then
                            strMsgCd = "W1150"
                            Exit Function
                        End If

                        '********* 電装ブロック＆給排気ブロック組み合わせチェック(7.11) *******************
                        Select Case strSeriesKata
                            Case "MN3Q0", "MT3Q0"
                                'MN3Q0,MT3Q0シリーズは対象外
                            Case Else
                                If strOptionT = "TX" Then
                                    bolFlag = False
                                    intLoop = intCI - 1
                                    Do While intLoop >= 0
                                        If arySelectInf(Siyou_01.Elect1 - 1)(intLoop) = "1" Or _
                                           arySelectInf(Siyou_01.Elect2 - 1)(intLoop) = "1" Then
                                            Exit Do
                                        End If
                                        For intRI2 As Integer = Siyou_01.Exhaust1 - 1 To Siyou_01.Exhaust4 - 1
                                            If arySelectInf(intRI2)(intLoop) = "1" And _
                                               strKataValues(intRI2).Contains("-C") Then
                                                bolFlag = True
                                                Exit Do
                                            End If
                                        Next
                                        intLoop = intLoop - 1
                                    Loop

                                    If Not bolFlag Then
                                        sbCoordinates.Append(Siyou_01.Elect1 & strComma & "0" & strPipe)
                                        sbCoordinates.Append(Siyou_01.Elect2 & strComma & "0")
                                        strMsg = sbCoordinates.ToString
                                        strMsgCd = "W1230"
                                        Exit Function
                                    End If
                                End If
                        End Select
                    End If
                Next

                '************ ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ ＆ 給排気ﾌﾞﾛｯｸ組み合わせチェック(7.32) ***************************
                If bolRegConb And _
                ((strSeriesKata <> "MN3EX0" Or strSeriesKata <> "MN4EX0") And objKtbnStrc.strcSelection.strOpSymbol(4) = "R" Or _
                 (strSeriesKata = "MN3EX0" Or strSeriesKata = "MN4EX0") And objKtbnStrc.strcSelection.strOpSymbol(2) = "R") Then
                    bolProc = False
                    intLoop = Siyou_01.Exhaust1 - 1
                    Do While intLoop < Siyou_01.Exhaust4
                        If arySelectInf(intLoop)(intCI) = "1" Then

                            bolProc = True
                            Exit Do
                        End If
                        intLoop = intLoop + 1
                    Loop
                    If bolProc Then
                        If fncConbCheck6(objKtbnStrc, intCI + 1, intColCnt - 1, 1, 2) Then
                            strCoordinates = CStr(intLoop + 1) & strComma & CStr(intCI + 1)
                            strMsg = strCoordinates
                            strMsgCd = "W1290"
                            Exit Function
                        End If
                        bolRegConb = False
                    End If
                End If
            Next

            'ミックスマニホールド特殊チェック(7.9.3)
            If ((strSeriesKata = "MN3E0" Or strSeriesKata = "MN4E0" Or _
               strSeriesKata = "MN3E00" Or strSeriesKata = "MN4E00") And _
               strOptionT IsNot Nothing And strOptionD IsNot Nothing And _
                objKtbnStrc.strcSelection.strOpSymbol(1) = "8" And objKtbnStrc.strcSelection.strOpSymbol(7) = "W") Or _
               ((strSeriesKata = "MN3EX0" Or strSeriesKata = "MN4EX0") And _
               strOptionT IsNot Nothing And strOptionD IsNot Nothing And _
                objKtbnStrc.strcSelection.strOpSymbol(5) = "W") Then

                intLoop = 0
                Do While intLoop < intColCnt
                    For intRI As Integer = Siyou_01.Valve1 - 1 To Siyou_01.Valve7 - 1
                        If arySelectInf(intRI)(intLoop) = "1" And _
                           arySelectInf(Siyou_01.Wiring - 1)(intLoop) = "0" Then

                            If strKataValues(intRI).Substring(0, 1) = "N" And _
                               (strKataValues(intRI).Contains("E010") Or _
                                strKataValues(intRI).Contains("E0110")) Then

                                Exit Do
                            End If
                        End If
                    Next
                    intLoop = intLoop + 1
                Loop
                '最終列まで条件を満たすものがない場合、エラー
                If intLoop = intColCnt Then
                    strMsgCd = "W1160"
                    Exit Function
                End If
            End If

            '********** 電磁弁付バルブブロック＆ＭＰ付バルブブロックの使用数チェック(7.10) *****
            bolFlag = False
            bolFlag2 = False
            bolFlag4 = False
            bolFlag5 = False
            bolFlag6 = False
            intElectSeq = 0
            'For intI As Integer = 0 To UBound(bolMixCon)
            '    bolMixCon(intI) = False
            'Next
            For intI As Integer = 0 To UBound(bolMixSwtch)
                bolMixSwtch(intI) = False
            Next

            For intRI As Integer = Siyou_01.Valve1 - 1 To Siyou_01.Valve7 - 1
                strKataban = Trim(strKataValues(intRI))
                If Len(strKataban) > 0 And _
                   Int(strUseValues(intRI)) > 0 Then

                    If (strSeriesKata = "MN3E0" Or strSeriesKata = "MN4E0" Or _
                       strSeriesKata = "MN3E00" Or strSeriesKata = "MN4E00") And _
                       (Trim(strSwitchPos).Length < 1) Then
                    ElseIf strSeriesKata = "MN3EX0" Or strSeriesKata = "MN4EX0" Or _
                           ((strSeriesKata <> "MN3EX0" Or strSeriesKata <> "MN4EX0") And _
                           (Left(strSwitchPos, 1) = "8" Or Left(strSwitchPos, 1) = "")) Then
                        If Left(strKataban, 4) = "N3E0" Or Left(strKataban, 4) = "N4E0" Then
                            If strKataban.Length < 7 Then
                            ElseIf Left(strKataban, 3) = "N3E" Then
                                Select Case strKataban.Substring(4, 3)
                                    Case "10-"
                                        bolMixSwtch(0) = True
                                        bolFlag = True
                                        bolFlag4 = True
                                        bolFlag5 = True
                                    Case "110"
                                        bolMixSwtch(1) = True
                                        bolFlag = True
                                        bolFlag4 = True
                                        bolFlag5 = True
                                    Case "20-"
                                        bolMixSwtch(2) = True
                                        bolFlag4 = True
                                        bolFlag5 = True
                                    Case "210"
                                        bolMixSwtch(3) = True
                                        bolFlag4 = True
                                        bolFlag5 = True
                                    Case "660"
                                        bolMixSwtch(4) = True
                                        bolFlag4 = True
                                        bolFlag5 = True
                                    Case "66S"
                                        bolMixSwtch(5) = True
                                        bolFlag4 = True
                                        bolFlag5 = True
                                    Case "670"
                                        bolMixSwtch(6) = True
                                        bolFlag4 = True
                                        bolFlag5 = True
                                    Case "67S"
                                        bolMixSwtch(7) = True
                                        bolFlag4 = True
                                        bolFlag5 = True
                                    Case "760"
                                        bolMixSwtch(8) = True
                                        bolFlag4 = True
                                        bolFlag5 = True
                                    Case "76S"
                                        bolMixSwtch(9) = True
                                        bolFlag4 = True
                                        bolFlag5 = True
                                    Case "770"
                                        bolMixSwtch(10) = True
                                        bolFlag4 = True
                                        bolFlag5 = True
                                    Case "77S"
                                        bolMixSwtch(11) = True
                                        bolFlag4 = True
                                        bolFlag5 = True
                                End Select
                            ElseIf Left(strKataban, 3) = "N4E" Then
                                Select Case strKataban.Substring(4, 3)
                                    Case "10-"
                                        bolMixSwtch(12) = True
                                        bolFlag = True  'ｼﾝｸﾞﾙｿﾚﾉｲﾄﾞ電磁弁指定チェック
                                        bolFlag2 = True '４ﾎﾟｰﾄ電磁弁指定チェック
                                        bolFlag4 = True '10mmタイプ指定チェック
                                    Case "20-"
                                        bolMixSwtch(13) = True
                                        bolFlag2 = True
                                        bolFlag4 = True
                                    Case "30-"
                                        bolMixSwtch(14) = True
                                        bolFlag2 = True
                                        bolFlag4 = True
                                    Case "40-"
                                        bolMixSwtch(15) = True
                                        bolFlag2 = True
                                        bolFlag4 = True
                                    Case "50-"
                                        bolMixSwtch(16) = True
                                        bolFlag2 = True
                                        bolFlag4 = True
                                End Select
                            End If
                        End If

                        If Left(strKataban, 5) = "N3E00" Or Left(strKataban, 5) = "N4E00" Then
                            If strKataban.Length < 7 Then
                            ElseIf Left(strKataban, 3) = "N3E" Then
                                Select Case strKataban.Substring(5, 3)
                                    Case "10-"
                                        bolMixSwtch(0) = True
                                        bolFlag = True
                                        bolFlag5 = True
                                        bolFlag6 = True
                                    Case "110"
                                        bolMixSwtch(1) = True
                                        bolFlag = True
                                        bolFlag5 = True
                                        bolFlag6 = True
                                    Case "20-"
                                        bolMixSwtch(2) = True
                                        bolFlag5 = True
                                        bolFlag6 = True
                                    Case "210"
                                        bolMixSwtch(3) = True
                                        bolFlag5 = True
                                        bolFlag6 = True
                                    Case "660"
                                        bolMixSwtch(4) = True
                                        bolFlag5 = True
                                        bolFlag6 = True
                                    Case "66S"
                                        bolMixSwtch(5) = True
                                        bolFlag5 = True
                                        bolFlag6 = True
                                    Case "670"
                                        bolMixSwtch(6) = True
                                        bolFlag5 = True
                                        bolFlag6 = True
                                    Case "67S"
                                        bolMixSwtch(7) = True
                                        bolFlag5 = True
                                        bolFlag6 = True
                                    Case "760"
                                        bolMixSwtch(8) = True
                                        bolFlag5 = True
                                        bolFlag6 = True
                                    Case "76S"
                                        bolMixSwtch(9) = True
                                        bolFlag5 = True
                                        bolFlag6 = True
                                    Case "770"
                                        bolMixSwtch(10) = True
                                        bolFlag5 = True
                                        bolFlag6 = True
                                    Case "77S"
                                        bolMixSwtch(11) = True
                                        bolFlag5 = True
                                        bolFlag6 = True
                                End Select
                            ElseIf Left(strKataban, 3) = "N4E" Then
                                Select Case strKataban.Substring(5, 3)
                                    Case "10-"
                                        bolMixSwtch(12) = True
                                        bolFlag = True
                                        bolFlag2 = True
                                        bolFlag6 = True
                                    Case "20-"
                                        bolMixSwtch(13) = True
                                        bolFlag2 = True
                                        bolFlag6 = True
                                    Case "30-"
                                        bolMixSwtch(14) = True
                                        bolFlag2 = True
                                        bolFlag6 = True
                                    Case "40-"
                                        bolMixSwtch(15) = True
                                        bolFlag2 = True
                                        bolFlag6 = True
                                    Case "50-"
                                        bolMixSwtch(16) = True
                                        bolFlag2 = True
                                        bolFlag6 = True
                                End Select
                            End If
                        End If
                    End If
                End If

                intElectSeq = intElectSeq + Int(strUseValues(intRI))
            Next

            For intRI As Integer = Siyou_01.Dummy1 - 1 To Siyou_01.Dummy2 - 1
                strKataban = strKataValues(intRI).Trim
                If Len(strKataban) > 0 And _
                   Int(strUseValues(intRI)) > 0 Then
                    bolMixSwtch(17) = True
                    intElectSeq = intElectSeq + Int(strUseValues(intRI))
                End If
            Next

            '最大連数値チェック
            If intElectSeq > Int(strMaxSeq) Then
                strMsgCd = "W1170"
                Exit Function

            ElseIf intElectSeq < Int(strMaxSeq) Then
                strMsgCd = "W1180"
                Exit Function

            End If

            '切り替え位置区分のチェック
            If Trim(strSwitchPos).Length < 1 Then
            ElseIf Left(strSwitchPos, 1) = "8" Then
                If (strSeriesKata = "MN4E0" Or strSeriesKata = "MN4E00") And _
                   Not bolFlag2 Then
                    strMsgCd = "W1200"
                    fncInpCheck2 = False
                    Exit Function
                End If

                If strSeriesKata = "MN3EX0" Or strSeriesKata = "MN4EX0" Then
                    If objKtbnStrc.strcSelection.strOpSymbol(5) = "W" And Not bolFlag Then
                        strMsgCd = "W1210"
                        Exit Function
                    End If
                Else
                    intCnt = 0
                    For intI As Integer = 0 To UBound(bolMixSwtch)
                        If bolMixSwtch(intI) Then
                            intCnt = intCnt + 1
                        End If
                    Next
                    If intCnt < 2 Then
                        strMsgCd = "W1190"
                        Exit Function
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(7) = "W" And Not bolFlag Then
                        strMsgCd = "W1210"
                        Exit Function
                    End If
                End If
            End If

            '接続口径のチェック
            If strConCaliber.StartsWith("CX") And
                   (strSeriesKata = "MN3E0" Or strSeriesKata = "MN4E0" Or _
                   strSeriesKata = "MN3E00" Or strSeriesKata = "MN4E00" Or _
                   strSeriesKata = "MN3Q0" Or strSeriesKata = "MT3Q0") Then
                intCnt = 0
                If Not SiyouBLL.fncMixBlockCheck(objKtbnStrc, Siyou_01.Valve1 - 1, Siyou_01.Valve7 - 1, strMsgCd) Then
                    Exit Function
                End If
            End If

            Select Case strSeriesKata
                Case "MN3EX0", "MN4EX0"
                    If bolFlag5 = False And bolFlag2 = True Then
                        If bolFlag4 = True And bolFlag6 = False Then
                            strMsgCd = "W8720" '"4ポート弁だけで10mmタイプのみの構成はMN4E0となります。"
                            Exit Function
                        ElseIf bolFlag4 = False And bolFlag6 = True Then
                            strMsgCd = "W8730" '"4ポート弁だけで7mmタイプのみの構成はMN4E00となります。"
                            Exit Function
                        End If
                    ElseIf bolFlag5 = True And bolFlag2 = False Then
                        If bolFlag4 = True And bolFlag6 = False Then
                            strMsgCd = "W8690" '"3ポート弁だけで10mmタイプのみの構成はMN3E0となります。"
                            Exit Function
                        ElseIf bolFlag4 = False And bolFlag6 = True Then
                            strMsgCd = "W8700" '"3ポート弁だけで7mmタイプのみの構成はMN3E00となります。"
                            Exit Function
                        ElseIf bolFlag4 = True And bolFlag6 = True And strSeriesKata = "MN4EX0" Then
                            strMsgCd = "W8710" '"3ポート弁だけで7mm,10mmミックスタイプの構成はMN3EX0となります。"
                            Exit Function
                        End If
                    Else
                        If bolFlag4 = True And bolFlag6 = False Then
                            strMsgCd = "W8740" '"10mmタイプのみの構成はMN4E0となります。"
                            Exit Function
                        ElseIf bolFlag4 = False And bolFlag6 = True Then
                            strMsgCd = "W8750" '"7mmタイプのみの構成はMN4E00となります。"
                            Exit Function
                        End If
                    End If
            End Select

            fncInpCheck2 = True

        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncInpCheck3
    '*【処理】
    '*   入力チェック
    '*   チェックに掛かっても処理を継続する
    '*【引数】
    '*   strKataValues  : 形番の選択値配列          strUseValues    : 使用数の入力値配列
    '*   arySelectInf   : 設置位置の選択値配列
    '********************************************************************************************
    Public Shared Function fncInpCheck3(objKtbnStrc As KHKtbnStrc, ByRef strMsg As String, ByRef strMsgCd As String) As Boolean

        Dim bolFlag As Boolean
        Dim intLoop As Integer = 0
        Dim strSeriesKata As String = String.Empty      '機種
        strSeriesKata = objKtbnStrc.strcSelection.strSeriesKataban

        fncInpCheck3 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            If (strSeriesKata.Trim <> "MN3EX0" Or strSeriesKata.Trim <> "MN4EX0") And objKtbnStrc.strcSelection.strOpSymbol(4) = "R" Or _
               (strSeriesKata.Trim = "MN3EX0" Or strSeriesKata.Trim = "MN4EX0") And objKtbnStrc.strcSelection.strOpSymbol(2) = "R" Then

                '************ ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ搭載時の警告メッセージ表示処理(7.37) ***************************
                bolFlag = False
                For intCI As Integer = 0 To intColCnt - 1
                    intLoop = Siyou_01.Valve1 - 1
                    Do While intLoop < Siyou_01.Valve7
                        If strSeriesKata.Trim = "MN3E0" Or strSeriesKata.Trim = "MN4E0" Or _
                           strSeriesKata.Trim = "MN3EX0" Or strSeriesKata.Trim = "MN4EX0" Then
                            If arySelectInf(intLoop)(intCI) = "1" And _
                               fncContaints(strKataValues(intLoop), "N3E066,N3E067,N3E076,N3E077") Then
                                bolFlag = True
                                Exit Do
                            End If
                        End If
                        If strSeriesKata.Trim = "MN3E00" Or strSeriesKata.Trim = "MN4E00" Or _
                           strSeriesKata.Trim = "MN3EX0" Or strSeriesKata.Trim = "MN4EX0" Then
                            If arySelectInf(intLoop)(intCI) = "1" And _
                               fncContaints(strKataValues(intLoop), "N3E0066,N3E0067,N3E0076,N3E0077") Then
                                bolFlag = True
                                Exit Do
                            End If
                        End If

                        intLoop = intLoop + 1
                    Loop
                    If bolFlag Then
                        strMsgCd = "W1300"
                        Exit Function
                    End If
                Next
            End If
            fncInpCheck3 = True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncContaints
    '*【処理】
    '*   指定された文字列の中に配列値が含まれているかチェック
    '*【引数】
    '*   strValue       : 文字列値          strKeys     : 配列値をカンマ区切りで連結した文字列
    '********************************************************************************************
    Public Shared Function fncContaints(ByVal strValue As String, ByVal strKeys As String) As Boolean
        Dim strKey() As String
        fncContaints = False
        Try
            strKey = strKeys.Split(strComma)
            For intI As Integer = 0 To UBound(strKey)
                If strValue.Contains(strKey(intI)) Then
                    fncContaints = True
                    Exit Function
                End If
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncSolenoidCnt1
    '*【処理】
    '*   確認処理(Ｔ＊)
    '*【引数】
    '*   intSt          : 処理開始位置                  intEd           : 処理終了位置
    '*   intAdd         : 処理方向(1:左→右 -1:右→左)  arySelectInf    : 設置位置の選択値配列
    '*   strKataValues  : 形番の選択値配列
    '********************************************************************************************
    Public Shared Function fncSolenoidCnt1(objKtbnStrc As KHKtbnStrc, ByVal intSt As Integer, ByVal intEd As Integer, _
                                     ByVal intAdd As Integer) As Integer
        Dim intCnt As Integer
        Dim intRI As Integer

        Try
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            intCnt = 0
            'ソレノイド点数カウント
            For intLoop As Integer = intSt To intEd Step intAdd
                intRI = 0
                Do While intRI < intPosRowCnt
                    If arySelectInf(intRI)(intLoop) = "1" Then
                        Select Case intRI + 1
                            Case Siyou_01.Elect1, Siyou_01.Elect2
                                Exit For
                            Case Siyou_01.Valve1, Siyou_01.Valve2, Siyou_01.Valve3, Siyou_01.Valve4, _
                                 Siyou_01.Valve5, Siyou_01.Valve6, Siyou_01.Valve7
                                intCnt = intCnt + 1
                            Case Siyou_01.Exhaust1, Siyou_01.Exhaust2, Siyou_01.Exhaust3, Siyou_01.Exhaust4
                                If strKataValues(intRI).Contains("-C") Then
                                    Exit For
                                End If
                            Case Siyou_01.EndL
                                If intAdd < 0 And strKataValues(intRI).Contains("-EL") Then
                                    Exit For
                                End If
                            Case Siyou_01.EndR
                                If intAdd > 0 Then
                                    Exit For
                                End If
                        End Select
                    End If
                    intRI = intRI + 1
                Loop
            Next

            fncSolenoidCnt1 = intCnt
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncSolenoidCnt2
    '*【処理】
    '*   確認処理(Ｄ＊)
    '*【引数】
    '*   intSt          : 処理開始位置                  intEd           : 処理終了位置
    '*   intAdd         : 処理方向(1:左→右 -1:右→左)  arySelectInf    : 設置位置の選択値配列
    '*   strKataValues  : 形番の選択値配列              intMax          : ソレノイドMAX値
    '********************************************************************************************
    Public Shared Function fncSolenoidCnt2(objKtbnStrc As KHKtbnStrc, _
                                           ByVal intSt As Integer, ByVal intEd As Integer, _
                                           ByVal intAdd As Integer, ByVal intMax As Integer) As Integer
        Dim intCnt As Integer
        Dim intRI As Integer
        Dim strKataban As String
        Dim WkKirikaeichi As String
        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban

        Try
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            intCnt = 0
            'ソレノイド点数カウント
            For intLoop As Integer = intSt To intEd Step intAdd
                intRI = 0
                Do While intRI < intPosRowCnt
                    If arySelectInf(intRI)(intLoop) = "1" Then
                        Select Case intRI + 1
                            Case Siyou_01.Elect1, Siyou_01.Elect2
                                Exit For

                            Case Siyou_01.Valve1, Siyou_01.Valve2, Siyou_01.Valve3, Siyou_01.Valve4, _
                                 Siyou_01.Valve5, Siyou_01.Valve6, Siyou_01.Valve7

                                If arySelectInf(Siyou_01.Wiring - 1)(intLoop) = "1" Then

                                    strKataban = strKataValues(intRI)

                                    If strKataban.Length < 5 Then
                                    Else
                                        If Mid(strKataban, 1, 5) = "N3E00" Or Mid(strKataban, 1, 5) = "N4E00" Then
                                            WkKirikaeichi = strKataban.Substring(5, 1)
                                        Else
                                            WkKirikaeichi = strKataban.Substring(4, 1)
                                        End If

                                        Select Case WkKirikaeichi
                                            Case "1"
                                                Select Case strSeriesKata.Trim
                                                    Case "MN3EX0", "MN4EX0"
                                                        If objKtbnStrc.strcSelection.strOpSymbol(5) = "W" Then
                                                            intCnt = intCnt + 2
                                                        Else
                                                            intCnt = intCnt + 1
                                                        End If
                                                    Case Else
                                                        If objKtbnStrc.strcSelection.strOpSymbol(7) = "W" Then
                                                            intCnt = intCnt + 2
                                                        Else
                                                            intCnt = intCnt + 1
                                                        End If
                                                End Select
                                            Case "-"
                                            Case Else
                                                intCnt = intCnt + 2
                                        End Select
                                    End If
                                End If
                                If intCnt > intMax Then
                                    Exit For
                                End If
                            Case Siyou_01.Dummy1, Siyou_01.Dummy2
                                strKataban = strKataValues(intRI).Trim
                                Select Case strKataban
                                    Case "N4E0-MPS"
                                        intCnt = intCnt + 1
                                    Case "N4E0-MPD"
                                        intCnt = intCnt + 2
                                    Case Else
                                End Select
                                If intCnt > intMax Then
                                    Exit For
                                End If

                            Case Siyou_01.Exhaust1, Siyou_01.Exhaust2, Siyou_01.Exhaust3, Siyou_01.Exhaust4
                                If strKataValues(intRI).Contains("-C") Then
                                    Exit For
                                End If

                            Case Siyou_01.EndL
                                If intAdd < 0 And strKataValues(intRI).Contains("-EL") Then
                                    Exit For
                                End If

                            Case Siyou_01.EndR
                                If intAdd > 0 Then
                                    Exit For
                                End If
                        End Select
                    End If
                    intRI = intRI + 1
                Loop
            Next

            fncSolenoidCnt2 = intCnt
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncSolenoidCnt3
    '*【処理】
    '*   確認処理(電装ブロックＬＲ)
    '*【引数】
    '*   intSt          : 処理開始位置                  intEd           : 処理終了位置
    '*   intAdd         : 処理方向(1:左→右 -1:右→左)  arySelectInf    : 設置位置の選択値配列
    '*   strKataValues  : 形番の選択値配列              intMax          : ソレノイドMAX値
    '********************************************************************************************
    Public Shared Function fncSolenoidCnt3(objKtbnStrc As KHKtbnStrc, _
                                          ByVal intSt As Integer, ByVal intEd As Integer, _
                                          ByVal intAdd As Integer, ByVal intMax As Integer) As Integer
        Dim intCnt As Integer
        Dim intRI As Integer
        Dim strKataban As String
        Dim WkKirikaeichi As String
        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban

        Try
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            intCnt = 0
            'ソレノイド点数カウント
            For intLoop As Integer = intSt To intEd Step intAdd
                intRI = 0
                Do While intRI < intPosRowCnt
                    If arySelectInf(intRI)(intLoop) = "1" Then
                        Select Case intRI + 1
                            Case Siyou_01.Elect1, Siyou_01.Elect2
                                If intAdd > 0 And strKataValues(intRI).Contains("TM") Then

                                    Exit For
                                End If

                            Case Siyou_01.Valve1, Siyou_01.Valve2, Siyou_01.Valve3, Siyou_01.Valve4, _
                                 Siyou_01.Valve5, Siyou_01.Valve6, Siyou_01.Valve7
                                If arySelectInf(Siyou_01.Wiring - 1)(intLoop) = "0" Then

                                    strKataban = strKataValues(intRI)

                                    If strKataban.Length < 5 Then
                                    Else
                                        If Mid(strKataban, 1, 5) = "N3E00" Or Mid(strKataban, 1, 5) = "N4E00" Then
                                            WkKirikaeichi = strKataban.Substring(5, 1)
                                        Else
                                            WkKirikaeichi = strKataban.Substring(4, 1)
                                        End If

                                        Select Case WkKirikaeichi
                                            Case "1"
                                                Select Case strSeriesKata.Trim
                                                    Case "MN3EX0", "MN4EX0"
                                                        If objKtbnStrc.strcSelection.strOpSymbol(5) = "W" Then
                                                            intCnt = intCnt + 2
                                                        Else
                                                            intCnt = intCnt + 1
                                                        End If
                                                    Case Else
                                                        If objKtbnStrc.strcSelection.strOpSymbol(7) = "W" Then
                                                            intCnt = intCnt + 2
                                                        Else
                                                            intCnt = intCnt + 1
                                                        End If
                                                End Select
                                            Case "-"
                                            Case Else
                                                intCnt = intCnt + 2
                                        End Select
                                    End If
                                End If
                                If intCnt > intMax Then
                                    Exit For
                                End If
                            Case Siyou_01.Dummy1, Siyou_01.Dummy2
                                strKataban = strKataValues(intRI).Trim
                                Select Case strKataban
                                    Case "N4E0-MPS"
                                        intCnt = intCnt + 1
                                    Case "N4E0-MPD"
                                        intCnt = intCnt + 2
                                    Case Else
                                End Select
                                If intCnt > intMax Then
                                    Exit For
                                End If

                            Case Siyou_01.Exhaust1, Siyou_01.Exhaust2, Siyou_01.Exhaust3, Siyou_01.Exhaust4
                                If strKataValues(intRI).Contains("-C") Then
                                    Exit For
                                End If

                            Case Siyou_01.EndL
                                If intAdd < 0 And strKataValues(intRI).Contains("-EL") Then
                                    Exit For
                                End If

                            Case Siyou_01.EndR
                                If intAdd > 0 Then
                                    Exit For
                                End If
                        End Select
                    End If
                    intRI = intRI + 1
                Loop
                If intAdd > 0 Then
                    If (arySelectInf(Siyou_01.Elect1)(intLoop) = "1" Or _
                             arySelectInf(Siyou_01.Elect2)(intLoop) = "1") And _
                            strKataValues(intRI).Contains("R") Then

                        Exit For
                    End If
                End If
            Next

            fncSolenoidCnt3 = intCnt
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   subConbCheck1
    '*【処理】
    '*   組み合わせチェック
    '*【引数】
    '*   bolFlag1       : 判定用フラグ              bolFlag2        : 判定用フラグ
    '*   bolFlag3       : 判定用フラグ              intStIdx        : 処理開始列No
    '*   arySelectInf   : 設置位置の選択値配列      strKataValues   : 形番の選択値配列
    '********************************************************************************************
    Public Shared Sub subConbCheck1(objKtbnStrc As KHKtbnStrc, ByRef bolFlag1 As Boolean, ByRef bolFlag2 As Boolean, _
                               ByRef bolFlag3 As Boolean, ByVal intStIdx As Integer)

        Try
            bolFlag1 = False
            bolFlag2 = False
            bolFlag3 = False
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            For intCI As Integer = intStIdx To 0 Step -1
                For intRI As Integer = 0 To intPosRowCnt - 1
                    If arySelectInf(intRI)(intCI) = "1" Then

                        Select Case intRI + 1
                            Case Siyou_01.Elect1, Siyou_01.Elect2
                                bolFlag3 = True
                                Exit Sub

                            Case Siyou_01.Valve1, Siyou_01.Valve2, Siyou_01.Valve3, Siyou_01.Valve4, _
                                 Siyou_01.Valve5, Siyou_01.Valve6, Siyou_01.Valve7
                                If arySelectInf(Siyou_01.Wiring - 1)(intCI) = "0" Then
                                    bolFlag1 = True
                                End If
                            Case Siyou_01.Exhaust1, Siyou_01.Exhaust2, Siyou_01.Exhaust3, Siyou_01.Exhaust4
                                If strKataValues(intRI).Contains("-C") Then
                                    bolFlag2 = True
                                End If
                        End Select
                    End If
                Next
            Next

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    '********************************************************************************************
    '*【関数名】
    '*   fncConbCheck2
    '*【処理】
    '*   組み合わせチェック
    '*【引数】
    '*   intStIdx       : 処理開始列No                   intEdIdx        : 処理終了列No
    '*   intAdd         : 1ループごとにIndexに加算する値 arySelectInf    : 設置位置の選択値配列
    '*   strKataValues  : 形番の選択値配列               strCoordinates  : エラーセルの座標
    '********************************************************************************************
    Public Shared Function fncConbCheck2(objKtbnStrc As KHKtbnStrc, ByVal intStIdx As Integer, ByVal intEdIdx As Integer, _
                                   ByVal intAdd As Integer, ByRef strCoordinates As String) As Boolean
        Dim bolFlag1 As Boolean
        Dim bolFlag2 As Boolean
        Dim intLoop As Integer

        Try
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            bolFlag1 = False
            bolFlag2 = False
            strCoordinates = ""

            For intCI As Integer = intStIdx To intEdIdx Step intAdd
                intLoop = 0
                Do While intLoop < arySelectInf.Count
                    If arySelectInf(intLoop)(intCI) = "1" Then

                        Select Case intLoop + 1
                            Case Siyou_01.Valve1, Siyou_01.Valve2, Siyou_01.Valve3, Siyou_01.Valve4, _
                                 Siyou_01.Valve5, Siyou_01.Valve6, Siyou_01.Valve7

                                bolFlag1 = True

                            Case Siyou_01.Exhaust1, Siyou_01.Exhaust2, Siyou_01.Exhaust3, Siyou_01.Exhaust4
                                If strKataValues(intLoop).Contains("-S") Then
                                    bolFlag2 = True
                                    strCoordinates = CStr(intLoop + 1) & strComma & CStr(intCI + 1)
                                    Exit For
                                End If

                            Case Siyou_01.EndR
                                If intAdd > 0 Then
                                    bolFlag2 = True
                                    Exit For
                                End If
                        End Select

                    End If
                    intLoop = intLoop + 1
                Loop
            Next

            If (Not bolFlag1) And bolFlag2 Then
                fncConbCheck2 = False
            Else
                fncConbCheck2 = True
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncConbCheck3
    '*【処理】
    '*   組み合わせチェック
    '*【引数】
    '*   intStIdx       : 処理開始列No              intDiv          : 処理区分
    '*   arySelectInf   : 設置位置の選択値配列      strKataValues   : 形番の選択値配列
    '*   strCoordinates : エラーセルの座標              
    '********************************************************************************************
    Public Shared Function fncConbCheck3(objKtbnStrc As KHKtbnStrc, ByVal intStIdx As Integer, ByVal intDiv As Integer, _
                                         ByRef strCoordinates As String) As Boolean

        Dim bolReturn As Boolean = True
        Dim intLoop As Integer
        Dim intLoop2 As Integer

        Try
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            If intDiv = 2 Then
                For intRI As Integer = Siyou_01.Regulat1 - 1 To Siyou_01.Regulat2 - 1
                    If arySelectInf(intRI)(intStIdx) = "1" And _
                       strKataValues(intRI).Contains("-FL") Then

                        fncConbCheck3 = True
                        Exit Function
                    End If
                Next
            End If
            For intCI As Integer = intStIdx To 0 Step -1
                intLoop = Siyou_01.Exhaust1 - 1
                Do While intLoop < Siyou_01.Exhaust4
                    If intDiv = 1 Then
                        If arySelectInf(intLoop)(intCI) = "1" And fncContaints(strKataValues(intLoop), "-SA") Then
                            If fncContaints(strKataValues(intLoop), "-Q-") Or _
                               fncContaints(strKataValues(intLoop), "-QK-") Or _
                               fncContaints(strKataValues(intLoop), "-QKZ-") Or _
                               fncContaints(strKataValues(intLoop), "-QX-") Or _
                               fncContaints(strKataValues(intLoop), "-QZ-") Or _
                               fncContaints(strKataValues(intLoop), "-QKX-") Then
                                fncConbCheck3 = True
                                Exit Function
                            End If
                            bolReturn = False
                            strCoordinates = CStr(intLoop + 1) & strComma & CStr(intCI + 1)
                            Exit For
                        End If
                    Else
                        If arySelectInf(intLoop)(intCI) = "1" Then
                            If strKataValues(intLoop).Contains("-S") And Not strKataValues(intLoop).Contains("-SA") Then

                                '右隣の列でﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸが選択されており、
                                '且つその行の形番の中に｢-FR｣が含まれていない場合、エラー
                                intLoop2 = Siyou_01.Regulat1 - 1
                                Do While intLoop2 < Siyou_01.Regulat2
                                    If arySelectInf(intLoop2)(intCI + 1) = "1" And _
                                       Not strKataValues(intLoop2).Contains("-FR") Then

                                        bolReturn = False
                                        strCoordinates = CStr(intLoop + 1) & strComma & CStr(intCI + 1)
                                        Exit For
                                    End If
                                    intLoop2 = intLoop2 + 1
                                Loop
                            End If

                            Exit For
                        End If
                    End If
                    intLoop = intLoop + 1
                Loop
            Next

            fncConbCheck3 = bolReturn
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   subConbCheck4
    '*【処理】
    '*   組み合わせチェック
    '*【引数】
    '*   bolFlag1       : 判定用フラグ              bolFlag2        : 判定用フラグ
    '*   bolFlag3       : 判定用フラグ              intStIdx        : 処理開始列No
    '*   intEdIdx       : 処理終了列No              intAdd          : 1ループごとにIndexに加算する値 
    '*   arySelectInf   : 設置位置の選択値配列      strKataValues   : 形番の選択値配列
    '*   strCoordinates : エラーセルの座標              
    '********************************************************************************************
    Public Shared Sub subConbCheck4(objKtbnStrc As KHKtbnStrc, ByRef bolFlag1 As Boolean, ByRef bolFlag2 As Boolean, _
                               ByRef bolFlag3 As Boolean, ByVal intStIdx As Integer, _
                               ByVal intEdIdx As Integer, ByVal intAdd As Integer, ByRef strCoordinates As String)

        Try
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            bolFlag1 = False
            bolFlag2 = False
            bolFlag3 = False

            For intCI As Integer = intStIdx To intEdIdx Step intAdd
                For intRI As Integer = Siyou_01.Exhaust1 - 1 To Siyou_01.Exhaust4 - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        If Not strKataValues(intRI).Contains("-SA") And _
                           fncContaints(strKataValues(intRI), "-Q-,-QK-,-QKX-,-QKZ-,-QX-") Then

                            bolFlag2 = True
                        ElseIf strKataValues(intRI).Contains("-SA") Then
                            If intAdd < 0 Or strKataValues(intRI).Contains("-QZ") Then
                                bolFlag1 = True
                                strCoordinates = CStr(intRI + 1) & strComma & CStr(intCI + 1)
                                Exit Sub
                            End If
                        End If
                    End If
                Next

                If intAdd > 0 Then
                    If arySelectInf(Siyou_01.EndR - 1)(intCI) = "1" Then
                        bolFlag3 = True
                        Exit Sub
                    End If

                    For intRI As Integer = Siyou_01.Elect1 - 1 To Siyou_01.Elect2 - 1
                        If arySelectInf(intRI)(intCI) = "1" Then
                            If strKataValues(intRI).Contains("R") Then
                                bolFlag3 = True
                                Exit Sub
                            End If
                        End If
                    Next
                End If

            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    '********************************************************************************************
    '*【関数名】
    '*   subConbCheck5
    '*【処理】
    '*   組み合わせチェック
    '*【引数】
    '*   bolFlag1       : 判定用フラグ              bolFlag2        : 判定用フラグ
    '*   bolFlag3       : 判定用フラグ              intStIdx        : 処理開始列No
    '*   intEdIdx       : 処理終了列No              intAdd          : 1ループごとにIndexに加算する値 
    '*   intDiv         : 処理区分                  arySelectInf    : 設置位置の選択値配列
    '*   strKataValues  : 形番の選択値配列          strCoordinates  : エラーセルの座標
    '********************************************************************************************
    Public Shared Sub subConbCheck5(objKtbnStrc As KHKtbnStrc, ByRef bolFlag1 As Boolean, ByRef bolFlag2 As Boolean, _
                              ByRef bolFlag3 As Boolean, ByVal intStIdx As Integer, _
                              ByVal intEdIdx As Integer, ByVal intAdd As Integer, _
                              ByVal intDiv As Integer, ByRef sbCoordinates As System.Text.StringBuilder)

        Try
            bolFlag1 = False
            bolFlag2 = False
            bolFlag3 = False

            For intCI As Integer = intStIdx To intEdIdx Step intAdd
                For intRI As Integer = Siyou_01.Valve1 - 1 To Siyou_01.Valve7 - 1
                    If objKtbnStrc.strcSelection.strPositionInfo(intRI)(intCI) = "1" Then
                        bolFlag1 = True
                        If intDiv = 1 Or intDiv = 3 Then
                            sbCoordinates.Append(CStr(intRI + 1) & strComma)
                            sbCoordinates.Append(CStr(intCI + 1) & strPipe)
                        End If
                    End If

                Next

                For intRI As Integer = Siyou_01.Exhaust1 - 1 To Siyou_01.Exhaust4 - 1
                    If objKtbnStrc.strcSelection.strPositionInfo(intRI)(intCI) = "1" Then
                        bolFlag2 = True
                    End If
                Next

                For intRI As Integer = Siyou_01.Regulat1 - 1 To Siyou_01.Regulat2 - 1
                    If objKtbnStrc.strcSelection.strPositionInfo(intRI)(intCI) = "1" Then
                        If intDiv = 2 Or _
                           intDiv = 4 Then
                            Exit Sub
                        ElseIf objKtbnStrc.strcSelection.strOptionKataban(intRI).Contains("-FR") Or _
                               objKtbnStrc.strcSelection.strOptionKataban(intRI).Contains("-FL") Then

                            bolFlag3 = True
                            If intDiv = 3 Then
                                Exit Sub
                            End If
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    '********************************************************************************************
    '*【関数名】
    '*   fncConbCheck6
    '*【処理】
    '*   組み合わせチェック
    '*【引数】
    '*   intStIdx       : 処理開始列No                   intEdIdx        : 処理終了列No
    '*   intAdd         : 1ループごとにIndexに加算する値 intProcDiv      : 処理区分
    '*   arySelectInf   : 設置位置の選択値配列           strKataValues   : 形番の選択値配列
    '********************************************************************************************
    Public Shared Function fncConbCheck6(objKtbnStrc As KHKtbnStrc, ByVal intStIdx As Integer, ByVal intEdIdx As Integer, _
                                   ByVal intAdd As Integer, ByVal intProcDiv As Integer) As Boolean

        Dim bolFlag1 As Boolean = False
        Dim bolFlag2 As Boolean = False
        Dim intLoop As Integer
        Dim strKey As String

        Try
            If intProcDiv = 1 Then
                If intAdd < 0 Then
                    strKey = "-FR"
                Else
                    strKey = "-FL"
                End If
            Else
                If intAdd < 0 Then
                    strKey = "-LR"
                Else
                    strKey = "-RL"
                End If
            End If

            For intCI As Integer = intStIdx To intEdIdx Step intAdd

                '7.29 または 7.30 の時のみ実行
                If intProcDiv = 1 Then
                    intLoop = Siyou_01.Exhaust1 - 1
                    Do While intLoop < Siyou_01.Exhaust4
                        If objKtbnStrc.strcSelection.strPositionInfo(intLoop)(intCI) = "1" Then
                            bolFlag1 = True
                            Exit For
                        End If
                        intLoop = intLoop + 1
                    Loop
                End If

                intLoop = Siyou_01.Regulat1 - 1
                Do While intLoop < Siyou_01.Regulat2
                    If objKtbnStrc.strcSelection.strPositionInfo(intLoop)(intCI) = "1" Then
                        If objKtbnStrc.strcSelection.strOptionKataban(intLoop).Contains(strKey) Then
                            bolFlag2 = True
                            Exit For
                        End If
                    End If
                    intLoop = intLoop + 1
                Loop
            Next

            If bolFlag1 Or bolFlag2 Then
                fncConbCheck6 = True
            Else
                fncConbCheck6 = False
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncCheckProc14
    '*【処理】
    '*   給排気ﾌﾞﾛｯｸ & 電装ﾌﾞﾛｯｸ組み合わせチェック
    '*【引数】
    '*   intColNo       : 処理開始列No              arySelectInf    : 設置位置の選択値配列
    '*   strKataValues  : 形番の選択値配列          intErrRow       : エラー行の行No
    '********************************************************************************************
    Public Shared Function fncCheckProc14(objKtbnStrc As KHKtbnStrc, ByVal intColNo As Integer, ByRef intErrRow As Integer) As Boolean
        Dim bolReturn As Boolean
        fncCheckProc14 = False
        Try
            bolReturn = True
            intErrRow = -1
            For intRI As Integer = 0 To objKtbnStrc.strcSelection.strPositionInfo.Count - 1
                If objKtbnStrc.strcSelection.strPositionInfo(intRI)(intColNo) = "1" Then
                    Select Case intRI + 1
                        Case Siyou_01.Elect1, Siyou_01.Elect2
                            If Not objKtbnStrc.strcSelection.strOptionKataban(intRI).Contains("R") And _
                                Not objKtbnStrc.strcSelection.strOptionKataban(intRI).Contains("TM") Then
                                intErrRow = intRI + 1
                                bolReturn = False
                            End If
                        Case Siyou_01.Exhaust1, Siyou_01.Exhaust2, Siyou_01.Exhaust3, Siyou_01.Exhaust4
                            If objKtbnStrc.strcSelection.strOptionKataban(intRI).Contains("-S") Then
                                intErrRow = intRI + 1
                                bolReturn = False
                            End If
                    End Select
                End If
            Next
            fncCheckProc14 = bolReturn
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncCheckProc22
    '*【処理】
    '*   給排気ﾌﾞﾛｯｸ & 電装ﾌﾞﾛｯｸ組み合わせチェック
    '*【引数】
    '*   intColNo       : 処理開始列No              arySelectInf    : 設置位置の選択値配列
    '*   strKataValues  : 形番の選択値配列
    '********************************************************************************************
    Public Shared Function fncCheckProc22(objKtbnStrc As KHKtbnStrc, ByVal intStIdx As Integer) As Boolean
        fncCheckProc22 = False
        Try
            For intCI As Integer = intStIdx To 0 Step -1
                For intRI As Integer = Siyou_01.Exhaust1 - 1 To Siyou_01.Exhaust4 - 1
                    If objKtbnStrc.strcSelection.strPositionInfo(intRI)(intCI) = "1" Then
                        If fncContaints(objKtbnStrc.strcSelection.strOptionKataban(intRI), "-QK-,-QKZ-,-QX-,-QKX-") Then
                            fncCheckProc22 = True
                            Exit Function
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   subCheckProc23
    '*【処理】
    '*   給排気ﾌﾞﾛｯｸ & ｴﾝﾄﾞﾌﾞﾛｯｸ組み合わせチェック
    '*【引数】
    '*   bolFlag1       : 判定用フラグ              bolFlag2        : 判定用フラグ
    '*   bolFlag3       : 判定用フラグ              intStIdx        : 処理開始列No
    '*   arySelectInf   : 設置位置の選択値配列      strKataValues   : 形番の選択値配列
    '*   strCoordinates : エラーセルの座標              
    '********************************************************************************************
    Public Shared Sub subCheckProc23(objKtbnStrc As KHKtbnStrc, ByRef bolFlag1 As Boolean, ByRef bolFlag2 As Boolean, _
                               ByRef bolFlag3 As Boolean, ByVal intStIdx As Integer, ByRef strCoordinates As String)

        Try
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            bolFlag1 = False
            bolFlag2 = False

            For intCI As Integer = intStIdx To intColCnt - 1
                For intRI As Integer = Siyou_01.Exhaust1 - 1 To Siyou_01.Exhaust4 - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        If strKataValues(intRI).Contains("-QZ-") Then
                            If strKataValues(intRI).Contains("-S") Then
                                bolFlag2 = True
                            End If
                        Else
                            bolFlag1 = True
                            strCoordinates = CStr(intRI + 1) & strComma & CStr(intCI + 1)
                            Exit Sub
                        End If
                    End If
                Next
                For intRI As Integer = Siyou_01.Regulat1 - 1 To Siyou_01.Regulat2 - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        If strKataValues(intRI).Contains("-RL") Or _
                           strKataValues(intRI).Contains("-FL") Then

                            bolFlag2 = True
                        End If
                    End If
                Next

                If arySelectInf(Siyou_01.EndR - 1)(intCI) = "1" Then
                    bolFlag3 = True
                    Exit Sub
                End If

                For intRI As Integer = Siyou_01.Elect1 - 1 To Siyou_01.Elect2 - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        If strKataValues(intRI).Contains("R") Then

                            bolFlag3 = True
                            Exit Sub
                        End If
                    End If
                Next
            Next

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Sub

    '********************************************************************************************
    '*【関数名】
    '*   fncCheckProc33
    '*【処理】
    '*   ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ & 給排気ﾌﾞﾛｯｸ組み合わせチェック
    '*【引数】
    '*   intStIdx       : 処理開始列No              arySelectInf    : 設置位置の選択値配列
    '*   strKataValues  : 形番の選択値配列          strCoordinates  : エラーセルの座標
    '********************************************************************************************
    Public Shared Function fncCheckProc33(objKtbnStrc As KHKtbnStrc, ByVal intStIdx As Integer, _
                                          ByRef strCoordinates As String) As Boolean

        Dim bolReturn As Boolean = True
        Dim bolFlag As Boolean = False
        Dim intLoop As Integer

        Try
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            For intCI As Integer = intStIdx To intColCnt - 1
                intLoop = Siyou_01.Exhaust1 - 1
                Do While intLoop < Siyou_01.Exhaust4
                    If arySelectInf(intLoop)(intCI) = "1" And strKataValues(intLoop).Contains("-S") And _
                       fncContaints(strKataValues(intLoop), "-QK-,-QKZ-,-QKX-,-QZ-") Then
                        bolFlag = True
                    End If
                    intLoop = intLoop + 1
                Loop
                Do While intLoop < Siyou_01.Regulat2
                    If arySelectInf(intLoop)(intCI) = "1" And _
                       fncContaints(strKataValues(intLoop), "RA-RL,RA-FL,RB-RL,RB-FL") Then
                        If Not bolFlag Then
                            bolReturn = False
                            strCoordinates = CStr(intLoop + 1) & strComma & CStr(intCI + 1)
                        End If
                        Exit For
                    End If
                    intLoop = intLoop + 1
                Loop

                If arySelectInf(Siyou_01.EndR - 1)(intCI) = "1" Then Exit For

                intLoop = Siyou_01.Elect1 - 1
                Do While intLoop < Siyou_01.Elect2
                    If arySelectInf(intLoop)(intCI) = "1" And strKataValues(intLoop).Contains("R") Then
                        Exit For
                    End If
                    intLoop = intLoop + 1
                Loop
            Next

            fncCheckProc33 = bolReturn
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncCheckProc34
    '*【処理】
    '*   ﾚｷﾞｭﾚｰﾀﾌﾞﾛｯｸ & 給排気ﾌﾞﾛｯｸ & ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ組み合わせチェック
    '*【引数】
    '*   intStIdx       : 処理開始列No              arySelectInf    : 設置位置の選択値配列
    '*   strKataValues  : 形番の選択値配列          strCoordinates  : エラーセルの座標
    '********************************************************************************************
    Public Shared Function fncCheckProc34(objKtbnStrc As KHKtbnStrc, ByVal intStIdx As Integer, _
                                          ByRef sbCoordinates As System.Text.StringBuilder) As Boolean

        Dim bolFlag1 As Boolean = False
        Dim bolFlag2 As Boolean = False
        Dim bolFlag3 As Boolean = False
        Dim intLoop As Integer
        fncCheckProc34 = False
        Try
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            For intCI As Integer = intStIdx To 0 Step -1
                If intCI = intStIdx Then
                    intLoop = Siyou_01.Valve1 - 1
                    Do While intLoop < Siyou_01.Valve7
                        If arySelectInf(intLoop)(intCI) = "1" Then
                            bolFlag1 = True
                            sbCoordinates.Append(CStr(intLoop + 1) & strComma)
                            sbCoordinates.Append(CStr(intCI + 1) & strPipe)
                        End If
                        intLoop = intLoop + 1
                    Loop
                End If

                intLoop = Siyou_01.Exhaust1 - 1
                Do While intLoop < Siyou_01.Exhaust4
                    If arySelectInf(intLoop)(intCI) = "1" And _
                       Not strKataValues(intLoop).Contains("-S") Then
                        bolFlag2 = True
                        Exit For
                    End If
                    intLoop = intLoop + 1
                Loop

                intLoop = Siyou_01.Regulat1 - 1
                Do While intLoop < Siyou_01.Regulat2
                    If arySelectInf(intLoop)(intCI) = "1" And _
                       fncContaints(strKataValues(intLoop), "-LR,-FR") Then
                        bolFlag3 = True
                        Exit For
                    End If
                    intLoop = intLoop + 1
                Loop
            Next

            If bolFlag1 And Not (bolFlag2 Or bolFlag3) Then
                fncCheckProc34 = False
            Else
                fncCheckProc34 = True
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   MN3Q0_Error
    '*【処理】
    '*   T**R配線最終端指示の指定位置チェック
    '*【引数】
    '*   EOR            : エラーチェックフラグ 
    '*   arySelectInf   :  
    '*   arySelectInf   : 
    '*   arySelectInf   : 
    '********************************************************************************************
    Public Shared Sub MN3Q0_Error(objKtbnStrc As KHKtbnStrc, ByRef EOR As Integer, _
                            ByRef LeftCnt As Integer, ByRef RightCnt As Integer)

        Dim TRPOINT As Integer = Nothing
        Dim i As Integer
        Dim j As Integer
        Dim ER(11) As Integer
        Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

        Try
            For intLoop As Integer = 0 To intColCnt - 1
                If arySelectInf(2)(intLoop) = "1" Then
                    TRPOINT = intLoop
                End If
            Next intLoop

            i = 0
            LeftCnt = 0
            RightCnt = 0

            For i = 3 To 9
                For intj As Integer = 0 To intColCnt - 1
                    If arySelectInf(i)(intj) = "1" And TRPOINT > intj Then LeftCnt = LeftCnt + 1 'TX用左加算処理
                    If arySelectInf(i)(intj) = "1" And TRPOINT <= intj Then RightCnt = RightCnt + 1 'TX用右加算処理
                    If arySelectInf(i)(intj) <> "1" And TRPOINT = intj Then ER(i) = 1 'エラー
                    If arySelectInf(i)(intj) = "1" And TRPOINT = intj Then ER(i) = 0 'エラー解除
                Next
            Next

            If TRPOINT >= 0 And TRPOINT <= intColCnt Then
                For i = 3 To 9
                    If ER(i) = 0 Then j = i
                Next i
                If j = 3 Or j = 4 Or j = 5 Or j = 6 Or j = 7 Or j = 8 Or j = 9 Then
                    EOR = 0
                Else
                    EOR = 1
                End If
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Sub
End Class
