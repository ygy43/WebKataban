Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_04
    Public Shared intPosRowCnt As Integer = 16      'RM1803032_スペーサ行追加
    Public Shared intColCnt As Integer = 40         'RM1803032_マニホールド連数拡張

    Public Shared Function fncInputChk(objKtbnStrc As KHKtbnStrc, HT_Option As Hashtable, dblStdNum As Double, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        fncInputChk = False
        Try
            '入力チェック
            If Not fncInpCheck2(objKtbnStrc, HT_Option, strMsg, strMsgCd) Then
                '画面作成
                Exit Function
            End If

            '入力チェック2
            If Not fncInpCheck1(objKtbnStrc, HT_Option, dblStdNum, strMsg, strMsgCd) Then
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
    '*   strKataValues  : 選択値(形番)配列          strCXAKataban   : 選択値(継手CXA)配列          
    '*   strCXBKataban  : 選択値(継手CXA)配列       strUseValues    : 使用数の入力値配列
    '*   arySelectInf   : 設置位置の選択値配列      dblDinRail      : Dinレール長さ
    '********************************************************************************************
    Public Shared Function fncInpCheck1(objKtbnStrc As KHKtbnStrc, HT_Option As Hashtable, dblStdNum As Double, _
                                        ByRef strMsg As String, ByRef strMsgCd As String) As Boolean

        Dim intPosCnt As Integer
        Dim intValPosCnt As Integer
        Dim intSpcPosCnt As Integer
        Dim intLoop As Integer = 0
        Dim intLoop2 As Integer = 0
        Dim sbCoordinates As New System.Text.StringBuilder
        Dim strCoordinates As String
        Dim intLeftEdge As Integer
        Dim intRightEdge As Integer
        Dim bolFlag As Boolean
        Dim strOptionD As String = HT_Option("strOptionD").ToString
        Dim strOptionX As String = HT_Option("strOptionX").ToString
        Dim strOptionG As String = HT_Option("strOptionG").ToString

        fncInpCheck1 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo
            Dim strCXAKataban() As String = objKtbnStrc.strcSelection.strCXAKataban
            Dim strCXBKataban() As String = objKtbnStrc.strcSelection.strCXBKataban

            '設置位置が選択されている行の形番が未選択の場合、エラー
            For intRI As Integer = 1 To strUseValues.Count
                If Int(strUseValues(intRI - 1)) > 0 And _
                   Len(Trim(strKataValues(intRI - 1))) = 0 And _
                   intRI <> Siyou_04.Rail Then
                    strMsgCd = "W1400"
                    Exit Function
                End If
            Next

            '未接続位置がある場合、エラー
            For intCI As Integer = 0 To intColCnt - 1
                intLoop = 0
                Do While intLoop < intPosRowCnt
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
                bolFlag = False
                For intRI As Integer = 0 To intPosRowCnt - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        '一番右側の選択列Ｎｏを取得
                        If intRightEdge = 0 Then
                            intRightEdge = intCI + 1
                        End If

                        bolFlag = True
                        Exit For
                    End If
                Next
                '中間に一つも選択されていない列がある場合、エラー
                If intRightEdge > 0 And Not bolFlag Then
                    strMsgCd = "W1020"
                    Exit Try
                End If
            Next

            'RM1803032_スペーサチェック変更
            '列ごとのチェック
            For intCI As Integer = 0 To intColCnt - 1

                '列ごとの選択数チェック
                intPosCnt = 0           '列全体の選択数
                intValPosCnt = 0        'バルブ＆マスキングプレートの選択数
                intSpcPosCnt = 0        'スペーサの選択数

                sbCoordinates = Nothing
                sbCoordinates = New System.Text.StringBuilder

                For intRI As Integer = 0 To intPosRowCnt - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        If intRI >= Siyou_04.Valve1 - 1 And _
                           intRI <= Siyou_04.MasPlate2 - 1 Then

                            'バルブ＆マスキングプレートが選択されている場合
                            intValPosCnt = intValPosCnt + 1

                        ElseIf intRI >= Siyou_04.Spacer1 - 1 And _
                               intRI <= Siyou_04.Spacer4 - 1 Then

                            'スペーサが選択されている場合
                            intSpcPosCnt = intSpcPosCnt + 1

                        End If

                        intPosCnt = intPosCnt + 1
                        sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)

                    End If
                Next

                '1つの列で３個以上選択されていたらエラー
                'バルブ＆マスキングプレートが同じ列に選択されていたらエラー
                'スペーサが同じ列に選択されていたらエラー
                If intPosCnt > 2 Or intValPosCnt > 1 Or intSpcPosCnt > 1 Then
                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                    strMsg = strCoordinates
                    strMsgCd = "W1390"
                    Exit Function
                End If
            Next

            'ﾌﾞﾗﾝｸﾌﾟﾗｸﾞ&ｻｲﾚﾝｻ、検査成績所の使用数チェック
            If Not SiyouBLL.fncOtherKataCheck(objKtbnStrc, Siyou_04.BlkPlug1, Siyou_04.Cable2, _
                         Siyou_04.Rail, strMsgCd) Then
                Exit Function
            End If

            'ﾊﾞﾙﾌﾞﾌﾞﾛｯｸ形番リスト重複チェック(形番＋継手CXA＋継手CXB)
            For intRI As Integer = Siyou_04.Valve1 - 1 To Siyou_04.Valve10 - 1
                For intRI2 As Integer = intRI + 1 To Siyou_04.Valve10 - 1
                    If Len(strKataValues(intRI)) = 0 Then
                    ElseIf ((strKataValues(intRI).Trim & strCXAKataban(intRI).Trim & strCXBKataban(intRI).Trim) = _
                           (strKataValues(intRI2).Trim & strCXAKataban(intRI2).Trim & strCXBKataban(intRI2).Trim)) And _
                           (strUseValues(intRI) > 0 And strUseValues(intRI2) > 0) Then
                        strMsgCd = "W1330"
                        Exit Function
                    End If
                Next
            Next

            'RM1803032_スペーサ行追加対応
            'スペーサ形番リスト重複チェック
            If Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_04.Spacer1, Siyou_04.Spacer4) Then
                strMsgCd = "W1330"
                Exit Function
            End If
            'If Len(strKataValues(Siyou_04.Spacer1 - 1)) = 0 Then
            'ElseIf strKataValues(Siyou_04.Spacer1 - 1) = strKataValues(Siyou_04.Spacer2 - 1) Then
            '    strMsgCd = "W1330"
            '    Exit Function
            'End If

            '取付レール長さチェック
            If strOptionD = "D" Then
                If strKataValues(Siyou_04.Rail - 1).ToString.Length <= 0 Then strKataValues(Siyou_04.Rail - 1) = 0
                If Not SiyouBLL.fncRailchk(strKataValues(Siyou_04.Rail - 1), CDbl(strUseValues(Siyou_04.Rail - 1)), dblStdNum, strMsgCd) Then
                    strMsg = Siyou_04.Rail & ",0"
                    Exit Function
                End If
            End If

            'G1,G2,X,X1 バルブブロックチェック
            If strOptionX = "X" Or strOptionX = "X1" _
                Or strOptionG = "G1" Or strOptionG = "G2" Then
                If Not SiyouBLL.fncBlockCheck2(strUseValues, strKataValues, Siyou_04.Elect1 - 1, Siyou_04.Elect10 - 1, strMsgCd) Then
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
    '*   入力チェック
    '*【引数】
    '*   strKataValues  : 選択値(形番)配列             strCXAKataban   : 選択値(継手CXA)配列
    '*   strCXBKataban  : 選択値(継手CXB)配列          strUseValues    : 使用数の入力値配列
    '*   arySelectInf   : 設置位置の選択値配列
    '********************************************************************************************
    Public Shared Function fncInpCheck2(objKtbnStrc As KHKtbnStrc, HT_Option As Hashtable, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        Dim intI As Integer
        Dim intRI As Integer
        Dim intCI As Integer
        Dim intSolCnt As Integer
        Dim intLoop As Integer
        Dim intLoop2 As Integer
        Dim intElectSeq As Integer
        Dim strKataban As String
        Dim sbCoordinates As New System.Text.StringBuilder
        Dim strCoordinates As String = ""
        Dim strOption As String = ""

        Dim bolProc As Boolean
        Dim bolCL As Boolean
        Dim bolMasPlate As Boolean
        Dim bolSpcrZ1 As Boolean
        Dim bolSpcrZ3 As Boolean
        Dim bolSpcrIS As Boolean
        Dim bolMixSwtch(11) As Boolean

        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban
        Dim strMaxSeq As String = HT_Option("strMaxSeq").ToString
        Dim strOptionH As String = HT_Option("strOptionH").ToString
        Dim strOptionZ1 As String = HT_Option("strOptionZ1")
        Dim strOptionZ2 As String = HT_Option("strOptionZ2")
        Dim strOptionZ3 As String = HT_Option("strOptionZ3")
        Dim strPortSize As String = HT_Option("strPortSize")
        Dim strElecConType As String = HT_Option("strElecConType")

        fncInpCheck2 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo
            Dim strCXAKataban() As String = objKtbnStrc.strcSelection.strCXAKataban
            Dim strCXBKataban() As String = objKtbnStrc.strcSelection.strCXBKataban

            '************ ソレノイド点数チェック(7.2) ************************************
            intSolCnt = 0
            'New4G対応 2017/01/06
            Dim strDensen As String = ""        '電線／省配線接続
            Dim strKeyKata As String = objKtbnStrc.strcSelection.strKeyKataban

            Select Case strKeyKata.Trim
                Case "R", "U", "S", "V"
                    strDensen = objKtbnStrc.strcSelection.strOpSymbol(5).ToString.Trim               '電線接続
                Case Else
                    strDensen = objKtbnStrc.strcSelection.strOpSymbol(4).ToString.Trim               '電線接続
            End Select
            'New4G対応 End

            'ソレノイドＭＡＸをセット
            'If Left(objKtbnStrc.strcSelection.strOpSymbol(4).ToString.Trim, 1) = "T" Then
            If Left(strDensen.ToString.Trim, 1) = "T" Then 'New4G対応 2017/01/06
                'ソレノイド点数を計算し、ソレノイドＭＡＸより多い場合はエラー
                For intCI = 0 To intColCnt - 1
                    For intRI = Siyou_04.Valve1 - 1 To Siyou_04.MasPlate2 - 1
                        If arySelectInf(intRI)(intCI) = "1" Then
                            Select Case intRI
                                Case Siyou_04.Valve1 - 1 To Siyou_04.Valve10 - 1
                                    Select Case Mid(strKataValues(intRI).Trim, 5, 1)
                                        Case "-"
                                        Case "1"
                                            intSolCnt = intSolCnt + 1
                                        Case Else
                                            intSolCnt = intSolCnt + 2
                                    End Select
                                Case Siyou_04.MasPlate1 - 1 To Siyou_04.MasPlate2 - 1
                                    If strKeyKata = "R" Or strKeyKata = "U" Or strKeyKata = "S" Or strKeyKata = "V" Then
                                        Select Case Mid(strKataValues(intRI).Trim, 8, 1)
                                            Case "S"
                                                intSolCnt = intSolCnt + 1
                                            Case "D"
                                                intSolCnt = intSolCnt + 2
                                        End Select
                                    Else
                                        Select Case Mid(strKataValues(intRI).Trim, 7, 1)
                                            Case "S"
                                                intSolCnt = intSolCnt + 1
                                            Case "D"
                                                intSolCnt = intSolCnt + 2
                                        End Select
                                    End If
                            End Select
                        End If
                    Next
                    If intSolCnt > KHKataban.fncGetMaxSol(objKtbnStrc.strcSelection.strOpSymbol, 4) Then
                        strMsgCd = "W1150"
                        Exit Function
                    End If
                Next
            End If

            Dim flgC8 As Boolean = False
            Dim flgC10 As Boolean = False
            Dim flgCX1 As Boolean = False
            Dim flgCX2 As Boolean = False
            '************ 継手CXチェック(7.3/7.4) ************************************
            For intI = Siyou_04.Valve1 To Siyou_04.MasPlate2
                If InStr(1, strKataValues(intI - 1).Trim, "-CX") <> 0 Then
                    If Int(strUseValues(intI - 1)) > 0 Then
                        If Len(strCXAKataban(intI - 1).Trim) = 0 Or Len(strCXBKataban(intI - 1).Trim) = 0 Then
                            '継手CXAと継手CXBに値が入っていない場合
                            strMsgCd = "W1450"
                            Exit Function
                        ElseIf strCXAKataban(intI - 1).Trim = strCXBKataban(intI - 1).Trim Then
                            '継手CXAと継手CXBに同じ値が入っている場合
                            strMsgCd = "W1460"
                            Exit Function
                        End If
                    End If
                End If
                If strSeriesKata = "M4GB1" Then
                    If InStr(1, strKataValues(intI - 1).Trim, "-C8") <> 0 Then
                        flgC8 = True

                    ElseIf InStr(1, strKataValues(intI - 1).Trim, "-CX") = 0 And strKataValues(intI - 1).Trim <> Nothing Then
                        flgCX1 = True
                    End If
                End If
                If strSeriesKata = "M4GB2" Then
                    If InStr(1, strKataValues(intI - 1).Trim, "-C10") <> 0 Then
                        flgC10 = True

                    ElseIf InStr(1, strKataValues(intI - 1).Trim, "-CX") = 0 And strKataValues(intI - 1).Trim <> Nothing Then
                        flgCX2 = True
                    End If
                End If
            Next

            If flgC8 And flgCX1 Then
                strMsgCd = "W9140"
                Exit Function
            End If

            If flgC10 And flgCX2 Then
                strMsgCd = "W9150"
                Exit Function
            End If

            '********** 電磁弁＆マスキングプレートの使用数/接続口径の使用数チェック(7.7) *****
            bolProc = True
            intElectSeq = 0
            For intI = 0 To UBound(bolMixSwtch)
                bolMixSwtch(intI) = False
            Next
            For intRI = Siyou_04.Valve1 - 1 To Siyou_04.MasPlate2 - 1
                strKataban = Trim(strKataValues(intRI))
                If Len(strKataban) > 0 And _
                   Int(strUseValues(intRI)) > 0 Then
                    If Trim(objKtbnStrc.strcSelection.strOpSymbol(1).ToString).Length < 1 Then
                    ElseIf Left(objKtbnStrc.strcSelection.strOpSymbol(1).ToString, 1) = "8" Then
                        If InStr(strKataban.Trim, "-MP") <> 0 Then      'MP,MPS,MPD は共に切換位置区分としては１種類の扱い
                            bolMixSwtch(11) = True
                        ElseIf Mid(strKataban.Trim, 1, 2) = "3G" Then
                            Select Case Mid(strKataban.Trim, 5, 2)
                                Case "19", "18"
                                    bolMixSwtch(0) = True
                                Case "11"
                                    bolMixSwtch(1) = True
                                Case "66"
                                    bolMixSwtch(2) = True
                                Case "67"
                                    bolMixSwtch(3) = True
                                Case "76"
                                    bolMixSwtch(4) = True
                                Case "77"
                                    bolMixSwtch(5) = True
                            End Select
                        ElseIf Mid(strKataban.Trim, 1, 2) = "4G" Then
                            Select Case Mid(strKataban.Trim, 5, 2)
                                Case "19", "18"
                                    bolMixSwtch(6) = True
                                Case "29", "28"
                                    bolMixSwtch(7) = True
                                Case "39", "38"
                                    bolMixSwtch(8) = True
                                Case "49", "48"
                                    bolMixSwtch(9) = True
                                Case "59", "58"
                                    bolMixSwtch(10) = True
                            End Select
                        End If
                    End If
                End If
                intElectSeq = intElectSeq + Int(strUseValues(intRI))
            Next

            '最大連数値チェック
            If intElectSeq > Int(strMaxSeq) Then
                strMsgCd = "W1170"
                Exit Function
            ElseIf intElectSeq < Int(strMaxSeq) Then
                strMsgCd = "W1180"
                Exit Function
            End If

            '切替位置区分ミックス時のチェック
            If Trim(objKtbnStrc.strcSelection.strOpSymbol(1).ToString).Length < 1 Then
            ElseIf Left(objKtbnStrc.strcSelection.strOpSymbol(1).ToString, 1) = "8" Then

                '切替位置区分がミックスの場合、電磁弁の切替位置区分が２種類以上でないとエラー
                intI = 0
                For intLoop = 0 To UBound(bolMixSwtch)
                    If bolMixSwtch(intLoop) Then
                        intI = intI + 1
                    End If
                Next
                If intI < 2 Then
                    strMsgCd = "W1190"
                    Exit Function
                End If

                'オプションＨの場合
                If strOptionH = "H" Then
                    If bolMixSwtch(0) = False And bolMixSwtch(1) = False And _
                       bolMixSwtch(2) = False And bolMixSwtch(3) = False And _
                       bolMixSwtch(4) = False And bolMixSwtch(5) = False And _
                       bolMixSwtch(6) = False And bolMixSwtch(7) = False And _
                       bolMixSwtch(8) = True And bolMixSwtch(9) = False And _
                       bolMixSwtch(10) = True And bolMixSwtch(11) = False Then
                        bolProc = False
                    ElseIf bolMixSwtch(0) = False And bolMixSwtch(1) = False And _
                           bolMixSwtch(2) = False And bolMixSwtch(3) = False And _
                           bolMixSwtch(4) = False And bolMixSwtch(5) = False And _
                           bolMixSwtch(6) = False And bolMixSwtch(7) = False And _
                           bolMixSwtch(8) = True And bolMixSwtch(9) = False And _
                           bolMixSwtch(10) = False And bolMixSwtch(11) = True Then
                        bolProc = False
                    ElseIf bolMixSwtch(0) = False And bolMixSwtch(1) = False And _
                           bolMixSwtch(2) = False And bolMixSwtch(3) = False And _
                           bolMixSwtch(4) = False And bolMixSwtch(5) = False And _
                           bolMixSwtch(6) = False And bolMixSwtch(7) = False And _
                           bolMixSwtch(8) = False And bolMixSwtch(9) = False And _
                           bolMixSwtch(10) = True And bolMixSwtch(11) = True Then
                        bolProc = False
                    ElseIf bolMixSwtch(0) = False And bolMixSwtch(1) = False And _
                           bolMixSwtch(2) = False And bolMixSwtch(3) = False And _
                           bolMixSwtch(4) = False And bolMixSwtch(5) = False And _
                           bolMixSwtch(6) = False And bolMixSwtch(7) = False And _
                           bolMixSwtch(8) = True And bolMixSwtch(9) = False And _
                           bolMixSwtch(10) = True And bolMixSwtch(11) = True Then
                        bolProc = False
                    ElseIf bolMixSwtch(0) = False And bolMixSwtch(1) = False And _
                           bolMixSwtch(2) = False And bolMixSwtch(3) = False And _
                           bolMixSwtch(4) = False And bolMixSwtch(5) = False And _
                           bolMixSwtch(6) = False And bolMixSwtch(7) = False And _
                           bolMixSwtch(8) = True And bolMixSwtch(9) = False And _
                           bolMixSwtch(10) = False And bolMixSwtch(11) = False Then
                        bolProc = False
                    ElseIf bolMixSwtch(0) = False And bolMixSwtch(1) = False And _
                           bolMixSwtch(2) = False And bolMixSwtch(3) = False And _
                           bolMixSwtch(4) = False And bolMixSwtch(5) = False And _
                           bolMixSwtch(6) = False And bolMixSwtch(7) = False And _
                           bolMixSwtch(8) = False And bolMixSwtch(9) = False And _
                           bolMixSwtch(10) = True And bolMixSwtch(11) = False Then
                        bolProc = False
                    ElseIf bolMixSwtch(0) = False And bolMixSwtch(1) = False And _
                           bolMixSwtch(2) = False And bolMixSwtch(3) = False And _
                           bolMixSwtch(4) = False And bolMixSwtch(5) = False And _
                           bolMixSwtch(6) = False And bolMixSwtch(7) = False And _
                           bolMixSwtch(8) = False And bolMixSwtch(9) = False And _
                           bolMixSwtch(10) = False And bolMixSwtch(11) = True Then
                        bolProc = False
                    End If
                    If bolProc = False Then
                        strMsgCd = "W1660"
                        Exit Function
                    End If
                End If
            End If

            '接続口径のチェック 2017/4/7 修正
            ' If objKtbnStrc.strcSelection.strOpSymbol(3).ToString.StartsWith("CX") Then
            If strPortSize.ToString.StartsWith("CX") Then
                If Not SiyouBLL.fncMixBlockCheck(objKtbnStrc, Siyou_04.Valve1 - 1, Siyou_04.MasPlate2 - 1, strMsgCd) Then
                    Exit Function
                End If
            End If

            '該当シリーズの場合、５ポート弁指定有無チェック
            Select Case strSeriesKata
                Case "M4GA1", "M4GA2", "M4GA3", "M4GB1", "M4GB2"
                    '切替位置区分がミックスの場合
                    If Trim(objKtbnStrc.strcSelection.strOpSymbol(1).ToString).Length < 1 Then
                    ElseIf Left(objKtbnStrc.strcSelection.strOpSymbol(1).ToString, 1) = "8" Then
                        Dim isOK As Boolean = False
                        '電磁弁の選択チェック
                        For i As Integer = Siyou_04.Valve1 - 1 To Siyou_04.Valve10 - 1
                            '5ポート弁(4GXXX)が選択されていること
                            If InStr(strKataValues(i), "4G") > 0 AndAlso strUseValues(i) > 0 Then
                                isOK = True
                                Exit For
                            End If
                        Next
                        '5ポート弁(4GXXX)の有無
                        If Not isOK Then
                            strMsgCd = "W0850"
                            Exit Function
                        End If
                    End If
            End Select

            '************ 電磁弁＆マスキングプレート継手チェック(7.9) ************************************
            Dim strCheckCX(strCXAKataban.Length) As String
            Dim strCheckCXCount(strCXAKataban.Length) As Double
            For intc As Integer = 0 To strCXAKataban.Length - 1
                strCheckCX(intc) = "-" & strCXAKataban(intc)
                strCheckCXCount(intc) = strUseValues(intc)
            Next
            For intc As Integer = 0 To strCXBKataban.Length - 1
                strCheckCX(intc) = "-" & strCXBKataban(intc)
                strCheckCXCount(intc) = strUseValues(intc)
            Next
            If Not SiyouBLL.fncBlockCheck(strUseValues, strKataValues, Siyou_04.Valve1 - 1, Siyou_04.MasPlate2 - 1, strMsgCd) Then
                Exit Function
            End If
            If Not SiyouBLL.fncBlockCheck(strCheckCXCount, strCheckCX, Siyou_04.Valve1 - 1, Siyou_04.MasPlate2 - 1, strMsgCd) Then
                Exit Function
            End If

            'W1
            Dim bolW1 As Boolean = False
            For i As Integer = Siyou_04.Valve1 - 1 To Siyou_04.Valve10 - 1
                If strKataValues(i).Length > 0 And CInt(strUseValues(i)) > 0 Then
                    Dim str As String = strKataValues(i).ToString.Substring(4, 1)
                    If str = "1" Then
                        bolW1 = True
                    End If
                End If
            Next

            If objKtbnStrc.strcSelection.strOpSymbol(6).ToString = "W1" And Not bolW1 Then
                'ダブル配線(W1)の時は、シングルソレノイドの選択が必要です。
                strMsgCd = "W9170"
                Exit Function
            End If

            '************ 二次電池対応チェック ************************************
            '* Ｍ４ＧＡ１で二次電池(Ｐ４)選択時
            '(ただし二次電池(Ｐ４)は、切換位置区分（８）選択かつ接続口径（Ｍ５）選択のみ選択可)
            '2017/4/7　修正
            '  If strSeriesKata = "M4GA1" AndAlso _
            '   objKtbnStrc.strcSelection.strOpSymbol(3).ToString = "M5" AndAlso _
            '   objKtbnStrc.strcSelection.strOpSymbol(13).ToString <> "" Then
            Select Case strKeyKata
                Case "R", "S", "U", "V"
                    If strSeriesKata = "M4GA1" AndAlso strPortSize = "M5" AndAlso objKtbnStrc.strcSelection.strOpSymbol(14).ToString <> "" Then
                        '* 仕様入力画面で電磁弁(3GA119-M5)または(3GA1119-M5)を選択していなければエラーとする
                        Dim isOK As Boolean = False
                        For i As Integer = 0 To 9
                            If (strKataValues(i).Equals("3GA119-M5") OrElse strKataValues(i).Equals("3GA1119-M5")) _
                            AndAlso strUseValues(i) > 0 Then
                                isOK = True
                            End If
                        Next
                        If Not isOK Then
                            '二次電池対応記号 'P4' は不要です。
                            strMsgCd = "W2750"
                            Exit Function
                        End If
                    End If
                Case Else
                    If strSeriesKata = "M4GA1" AndAlso strPortSize = "M5" AndAlso objKtbnStrc.strcSelection.strOpSymbol(13).ToString <> "" Then
                        '* 仕様入力画面で電磁弁(3GA119-M5)または(3GA1119-M5)を選択していなければエラーとする
                        Dim isOK As Boolean = False
                        For i As Integer = 0 To 9
                            If (strKataValues(i).Equals("3GA119-M5") OrElse strKataValues(i).Equals("3GA1119-M5")) _
                            AndAlso strUseValues(i) > 0 Then
                                isOK = True
                            End If
                        Next
                        If Not isOK Then
                            '二次電池対応記号 'P4' は不要です。
                            strMsgCd = "W2750"
                            Exit Function
                        End If
                    End If
            End Select

            'RM1803032_スペーサ行追加対応
            '************ スペーサ＆電磁弁＆マスキングプレート組合せチェック(7.8/7.11) ************************************
            bolSpcrZ1 = False
            bolSpcrZ3 = False
            bolSpcrIS = False

            sbCoordinates = Nothing
            sbCoordinates = New System.Text.StringBuilder

            For intCI = 0 To intColCnt - 1

                intLoop = Siyou_04.Spacer1 - 1

                Do While intLoop < Siyou_04.Spacer4
                    bolProc = False
                    bolCL = True
                    bolMasPlate = True

                    If arySelectInf(intLoop)(intCI) = "1" Then

                        If InStr(strKataValues(intLoop), "-P") Then
                            bolSpcrZ1 = True
                        ElseIf InStr(strKataValues(intLoop), "-R") Then
                            bolSpcrZ3 = True
                        ElseIf InStr(strKataValues(intLoop), "-IS") Then
                            bolSpcrIS = True
                        End If

                        '電磁弁＆マスキングプレートが同じ位置に設置されていなければエラー
                        '電線接続が"T*"の場合、スペーサを指定している列の電磁弁＆マスキングプレートにワンタッチ継手("-CL"を含む)を選択していたらエラー
                        'マスキングプレートと同じ位置に給気スペーサ("-P"を含む)を選択していたらエラー
                        For intLoop2 = Siyou_04.Valve1 - 1 To Siyou_04.MasPlate2 - 1
                            If arySelectInf(intLoop2)(intCI) = "1" Then

                                bolProc = True
                                '2017/4/7　修正
                                '  If (bolSpcrZ1 = True Or bolSpcrIS = True) And Left(objKtbnStrc.strcSelection.strOpSymbol(4).ToString, 1) = "T" Then
                                If (bolSpcrZ1 = True Or bolSpcrIS = True) And Left(strElecConType, 1) = "T" Then
                                    If InStr(strKataValues(intLoop2), "-CL") = 0 And _
                                       InStr(strCXAKataban(intLoop2), "CL") = 0 And _
                                       InStr(strCXBKataban(intLoop2), "CL") = 0 Then
                                    Else
                                        bolCL = False
                                        sbCoordinates.Append(CStr(intLoop + 1) & strComma & CStr(intCI + 1) & strPipe & CStr(intLoop2 + 1) & strComma & CStr(intCI + 1) & strPipe)
                                    End If
                                End If
                                If (bolSpcrZ1 = True Or bolSpcrZ3 = True Or bolSpcrIS = True) And (intLoop2 = Siyou_04.MasPlate2 - 1 Or intLoop2 = Siyou_04.MasPlate1 - 1) Then
                                    sbCoordinates.Append(CStr(intLoop + 1) & strComma & CStr(intCI + 1) & strPipe & CStr(intLoop2 + 1) & strComma & CStr(intCI + 1) & strPipe)
                                    bolMasPlate = False
                                End If
                            End If
                        Next

                        If Not bolProc Then
                            sbCoordinates.Append(CStr(intLoop + 1) & strComma & CStr(intCI + 1) & strPipe)
                            strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                            strMsg = strCoordinates
                            strMsgCd = "W4010"
                            Exit Function
                        ElseIf Not bolCL Then
                            strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                            strMsg = strCoordinates
                            strMsgCd = "W4140"
                            Exit Function
                        ElseIf Not bolMasPlate Then
                            strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                            strMsg = strCoordinates
                            strMsgCd = "W4020"
                            Exit Function
                        End If
                    End If

                    intLoop = intLoop + 1
                Loop
            Next

            '引当画面のオプションで「Z1」「Z2」「Z3」を含んでいる場合は、スペーサを選択しないとエラー
            Select Case strSeriesKata
                Case "M4GA4", "M4GB4", "M4GD4", "M4GE4"
                    If strOptionZ1 = "Z1" Then
                        If Not bolSpcrZ1 Then
                            strMsgCd = "W4030"
                            Exit Function
                        End If
                    End If

                    If strOptionZ3 = "Z3" Then
                        If Not bolSpcrZ3 Then
                            strMsgCd = "W4040"
                            Exit Function
                        End If
                    End If
                Case Else
                    If strOptionZ1 = "Z1" Or strOptionZ2 = "Z2" Then
                        If Not bolSpcrZ1 And Not bolSpcrZ3 And Not bolSpcrIS Then
                            strMsgCd = "W4130"
                            Exit Function
                        End If
                    End If
                    If strOptionZ1 = "Z1" And strOptionZ2 = "Z2" Then
                        If Not bolSpcrZ1 Or (Not bolSpcrZ3 And Not bolSpcrIS) Then
                            strMsgCd = "W4180"
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

End Class
