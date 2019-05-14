Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_03

    Public Shared intPosRowCnt As Integer = 15
    Public Shared intColCnt As Integer = 20

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
    '*   fncInpCheck2
    '*【処理】
    '*   入力チェック
    '*【引数】
    '*   strKataValues  : 形番の選択値配列          strUseValues    : 使用数の入力値配列
    '*   arySelectInf   : 設置位置の選択値配列
    '********************************************************************************************
    Public Shared Function fncInpCheck2(objKtbnStrc As KHKtbnStrc, ByRef dblStdNum As Double, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        Dim sbCoordinates As New System.Text.StringBuilder
        fncInpCheck2 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban

            '形番選択チェック
            For intRI As Integer = 0 To intPosRowCnt - 1
                If strKataValues(intRI) = "" And strUseValues(intRI) > 0 Then
                    sbCoordinates.Append(CStr(intRI + 1) & strComma & "0")
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W1400"
                    Exit Function
                End If
            Next

            '電磁弁重複チェック
            For intRI As Integer = Siyou_03.Elect1 To Siyou_03.Elect14 - 1
                If strKataValues(intRI - 1).Length > 0 And InStr(strKataValues(intRI - 1), "-CX") = 0 Then
                    For intRI2 As Integer = intRI + 1 To Siyou_03.Elect14
                        If strKataValues(intRI2 - 1).Length > 0 And _
                           strKataValues(intRI - 1) = strKataValues(intRI2 - 1) Then
                            strMsgCd = "W1330"
                            Exit Function
                        End If
                    Next
                End If
            Next

            '取付レール長さ入力値チェック
            If strKataValues(Siyou_03.Rail - 1).ToString.Length <= 0 Then strKataValues(Siyou_03.Rail - 1) = 0
            If Not SiyouBLL.fncRailchk(strKataValues(Siyou_03.Rail - 1), CDbl(strUseValues(Siyou_03.Rail - 1)), dblStdNum, strMsgCd) Then
                strMsg = Siyou_03.Rail & ",0"
                Exit Function
            End If

            '添付部品形番選択チェック
            If Not SiyouBLL.fncOtherKataCheck(objKtbnStrc, Siyou_03.Plug1, Siyou_03.Cable2, Siyou_03.Rail, strMsgCd) Then
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
    '*   fncInpCheck1
    '*【処理】
    '*   入力チェック
    '*【引数】
    '*   strKataValues  : 形番の選択値配列          strUseValues    : 使用数の入力値配列
    '*   arySelectInf   : 設置位置の選択値配列      dblStdNum       : ﾏﾆﾎｰﾙﾄﾞ長さ基数
    '********************************************************************************************
    Public Shared Function fncInpCheck1(objKtbnStrc As KHKtbnStrc, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        Dim sbCoordinates As New System.Text.StringBuilder
        Dim intSolCnt As Integer
        Dim intColR As Integer
        Dim bolChkFlag As Boolean = False
        Dim bolMixCon(2) As Boolean
        Dim intMixSwtchCnt As Integer = 0
        Dim intMixConCnt As Integer = 0
        Dim intElectSeq As Integer
        Dim bolElect4Port As Boolean
        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban
        Dim strKeyKata As String = objKtbnStrc.strcSelection.strKeyKataban

        fncInpCheck1 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo
            Dim strCXAKataban() As String = objKtbnStrc.strcSelection.strCXAKataban
            Dim strCXBKataban() As String = objKtbnStrc.strcSelection.strCXBKataban

            '最右部列取得
            intColR = 0
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

            '全未チェックエラー
            If intColR = 0 Then
                strMsgCd = "W1030"
                Exit Function
            End If

            '列連続チェックエラー
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

            'ソレノイド点数チェック
            If Left(objKtbnStrc.strcSelection.strOpSymbol(9), 1) = "T" Then
                intSolCnt = 0
                For intCI As Integer = 1 To intColR
                    '１～１４行目：電磁弁
                    For intRI As Integer = Siyou_03.Elect1 To Siyou_03.Elect14
                        If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                            If strKataValues(intRI - 1).Substring(4, 1) = "1" Then
                                intSolCnt = intSolCnt + 1
                            ElseIf strKataValues(intRI - 1).Substring(4, 1) <> "-" Then
                                intSolCnt = intSolCnt + 2
                            End If
                        End If
                    Next

                    '１５行目：マスキングプレート
                    If arySelectInf(Siyou_03.Masking - 1)(intCI - 1) = "1" Then
                        intSolCnt = intSolCnt + 1
                    End If

                    If intSolCnt > KHKataban.fncGetMaxSol(objKtbnStrc.strcSelection.strOpSymbol, 3) Then
                        strMsgCd = "W1150"
                        Exit Function
                    End If
                Next
            End If

            ''********** 電磁弁・マスキングプレートチェック ******************************
            '１～１４行目：電磁弁
            For intI As Integer = 0 To UBound(bolMixCon)
                bolMixCon(intI) = False
            Next
            intElectSeq = 0
            bolElect4Port = False

            Dim ListMixSwtch As New ArrayList
            For intRI As Integer = Siyou_03.Elect1 - 1 To Siyou_03.Masking - 1
                '形番要素が選択されている場合
                If strKataValues(intRI).Length > 0 Then
                    'CXチェック(電磁弁)
                    If intRI < Siyou_03.Masking - 1 Then
                        If InStr(strKataValues(intRI), "-CX") > 0 Then
                            If CInt(strUseValues(intRI)) > 0 Then
                                'CXA・CXB未選択
                                If strCXAKataban(intRI) = "" Or strCXBKataban(intRI) = "" Then
                                    strMsgCd = "W1450"
                                    Exit Function
                                End If

                                '重複チェック
                                For intRI2 As Integer = Siyou_03.Elect1 To Siyou_03.Elect14 - 1
                                    If intRI <> intRI2 Then
                                        If strKataValues(intRI) = strKataValues(intRI2) And _
                                         ((strCXAKataban(intRI) = strCXAKataban(intRI2) And strCXBKataban(intRI) = strCXBKataban(intRI2)) Or _
                                          (strCXAKataban(intRI) = strCXBKataban(intRI2) And strCXBKataban(intRI) = strCXAKataban(intRI2))) Then
                                            strMsgCd = "W1330"
                                            Exit Function
                                        End If
                                    End If
                                Next
                            End If

                            'CXA・CXB同一チェック
                            If strCXAKataban(intRI) <> "" And strCXAKataban(intRI) = strCXBKataban(intRI) Then
                                strMsgCd = "W1460"
                                Exit Function
                            End If
                        End If

                        If strKataValues(intRI) = "" And CInt(strUseValues(intRI)) > 0 Then
                            strMsgCd = "W1400"
                            Exit Function
                        End If
                    Else
                        'CXチェック(マスキングプレート)
                        'CXA・CXB同一チェック
                        If strCXAKataban(intRI) <> "" And strCXAKataban(intRI) = strCXBKataban(intRI) Then
                            strMsgCd = "W1460"
                            Exit Function
                        End If
                        '未選択チェック
                        If InStr(strKataValues(intRI), "-CX") > 0 Then '2009/03/18 T.Y 不具合修正
                            If CInt(strUseValues(Siyou_03.Masking - 1)) > 0 Then
                                If strSeriesKata = "M" And (strKeyKata = "3" Or strKeyKata = "4") Then
                                    If strCXAKataban(Siyou_03.Masking - 1) = "" Or strCXBKataban(Siyou_03.Masking - 1) = "" Then
                                        strMsgCd = "W1450"
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If

                    If Left(objKtbnStrc.strcSelection.strOpSymbol(4).ToString, 1) = "8" Then
                        'ミックスチェック(切替位置区分)セット
                        If CInt(strUseValues(intRI)) > 0 Then
                            Dim strKey As String = Mid(strKataValues(intRI).ToString.PadRight(7, " "), 5, 2)
                            Select Case strKey
                                Case "19"
                                    Select Case Left(strKataValues(intRI), 2)
                                        Case "3S", "4S"
                                            If Not ListMixSwtch.Contains(strKey & "," & Left(strKataValues(intRI), 2)) Then
                                                ListMixSwtch.Add(strKey & "," & Left(strKataValues(intRI), 2))
                                            End If
                                            If Left(strKataValues(intRI), 2) = "4S" Then bolElect4Port = True
                                    End Select
                                Case "11"
                                    If strKataValues(intRI).ToString.StartsWith("3S") Then
                                        If Not ListMixSwtch.Contains(strKey & "," & Left(strKataValues(intRI), 2)) Then
                                            ListMixSwtch.Add(strKey & "," & Left(strKataValues(intRI), 2))
                                        End If
                                    End If
                                Case "29", "39", "49", "59"
                                    If strKataValues(intRI).ToString.StartsWith("4S") Then
                                        If Not ListMixSwtch.Contains(strKey & "," & Left(strKataValues(intRI), 2)) Then
                                            ListMixSwtch.Add(strKey & "," & Left(strKataValues(intRI), 2))
                                        End If
                                        bolElect4Port = True
                                    End If
                                Case Else
                                    If strKataValues(intRI).Trim = "MP" Then
                                        If Not ListMixSwtch.Contains(strKataValues(intRI).Trim) Then
                                            ListMixSwtch.Add(strKataValues(intRI).Trim)
                                        End If
                                    End If
                            End Select
                        End If
                    End If
                    'ミックスチェック(接続口径)セット
                    If objKtbnStrc.strcSelection.strOpSymbol(6).ToString.StartsWith("CX") Then
                        If CInt(strUseValues(intRI)) > 0 Then
                            If InStr(strKataValues(intRI), "-C4") > 0 Then
                                bolMixCon(1) = True
                            End If
                            If InStr(strKataValues(intRI), "-C6") > 0 Then
                                bolMixCon(2) = True
                            End If
                            If InStr(strKataValues(intRI), "-CX") > 0 Then
                                bolMixCon(1) = True
                                bolMixCon(2) = True
                            End If
                        End If
                    End If

                    '電磁弁連数値セット
                    If CInt(strUseValues(intRI)) > 0 Then
                        intElectSeq = intElectSeq + Int(strUseValues(intRI))
                    End If
                End If
            Next

            '電磁弁エラーチェック
            If intElectSeq > objKtbnStrc.strcSelection.strOpSymbol(10).ToString Then
                strMsgCd = "W1170"
                Exit Function
            End If

            If intElectSeq < objKtbnStrc.strcSelection.strOpSymbol(10).ToString Then
                strMsgCd = "W1180"
                Exit Function
            End If

            If Left(objKtbnStrc.strcSelection.strOpSymbol(4).ToString, 1) = "8" Then
                If ListMixSwtch.Count <= 1 Then
                    strMsgCd = "W1190"
                    Exit Function
                End If
            End If

            If Left(objKtbnStrc.strcSelection.strOpSymbol(4).ToString, 1) = "8" Then
                If strSeriesKata = "M" And (strKeyKata = "1" Or strKeyKata = "2") Then
                    If objKtbnStrc.strcSelection.strOpSymbol(2) = "4" And _
                        objKtbnStrc.strcSelection.strOpSymbol(3) = "SA1" Then
                        If bolElect4Port = False Then
                            strMsgCd = "W1470"
                            Exit Function
                        End If
                    End If
                End If
            End If

            If objKtbnStrc.strcSelection.strOpSymbol(6).ToString.StartsWith("CX") Then
                For intI As Integer = 0 To UBound(bolMixCon)
                    If bolMixCon(intI) Then
                        intMixConCnt = intMixConCnt + 1
                    End If
                Next
                If intMixConCnt <= 1 Then
                    strMsgCd = "W1480"
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
End Class
