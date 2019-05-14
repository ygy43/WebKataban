Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_11

    Public Shared intPosRowCnt As Integer = 15
    Public Shared intColCnt As Integer = 20

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
        Dim intColR As Integer = 0
        Dim bolChkFlag As Boolean = False
        Dim intChargeAirCnt As Integer          '集中給気ブロック数
        Dim intChargeAirAPSCnt As Integer       'APS付集中給気ブロック数
        Dim intTotalChargeAirCnt As Integer     '集中給気ブロック選択数計
        Dim bolRegA As Boolean = False          'レギュレータ(*500A*)タイプチェック
        Dim bolRegB As Boolean = False          'レギュレータ(*500B*)タイプチェック
        Dim intSelCnt As Integer = 0            'レギュレータ・サブベース選択数
        Dim intSubBaseCol As Integer = 0        'MP付サブベース選択列
        Dim dblRailLen As Double                '取付レール長さ
        Dim bolMixFlag As Boolean = False       'MIX構成判定フラグ
        Dim intChargePos As Integer
        Dim bolRegL As Boolean = False
        Dim bolRegR As Boolean = False
        Dim intRegCnt As Integer = 0            'レギュレータ選択数
        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban

        fncInputChk = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            '------ 既存システムで、コントロール上の値変更時にチェックしていた内容 ----------
            '9.1 形番選択チェック
            For intRI As Integer = 1 To intPosRowCnt
                If Trim(strKataValues(intRI - 1)).Length = 0 Then
                    For intCI As Integer = 1 To intColCnt
                        If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                            sbCoordinates.Append(CStr(intRI) & strComma & "0")
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W1400"
                            Exit Function
                        End If
                    Next
                End If
            Next

            '9.2 形番要素重複チェック
            If Not SiyouBLL.fncDblCheck(objKtbnStrc, Siyou_11.Regulator1, Siyou_11.Regulator10) Then
                strMsgCd = "W1330"
                Exit Function
            End If

            '9.3 ブランクプラグ個数チェック
            If Not SiyouBLL.fncOtherKataCheck(objKtbnStrc, Siyou_11.Plug1, Siyou_11.Plug3, _
                                     0, strMsgCd, 100) Then
                Exit Function
            End If

            '----- 入力内容チェック ----------
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

            '8.1 未接続位置チェック
            '8.1.1.2 一つも未選択の場合、エラー
            If intColR = 0 Then
                strMsgCd = "W1030"
                Exit Function
            End If
            '8.1.1.1 最右列まで連続チェックされていない場合、エラー
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
            '8.2 エンドブロックLチェック(１行目)
            '8.2.1 必須選択チェック
            If CInt(strUseValues(Siyou_11.EndL - 1)) = 0 Then
                sbCoordinates.Append(CStr(Siyou_11.EndL) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W1620"
                Exit Function
            End If
            '8.2.2 複数選択チェック
            If CInt(strUseValues(Siyou_11.EndL - 1)) > 1 Then
                sbCoordinates.Append(CStr(Siyou_11.EndL) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W1100"
                Exit Function
            End If
            '8.3 エンドブロックRチェック(１５行目)
            '8.3.1 必須選択チェック
            If CInt(strUseValues(Siyou_11.EndR - 1)) = 0 Then
                sbCoordinates.Append(CStr(Siyou_11.EndR) & strComma & CStr(intColR + 1))
                strMsg = sbCoordinates.ToString
                strMsgCd = "W1650"
                Exit Function
            End If
            '8.3.2 複数選択チェック
            If CInt(strUseValues(Siyou_11.EndR - 1)) > 1 Then
                sbCoordinates.Append(CStr(Siyou_11.EndR) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W1100"
                Exit Function
            End If
            '8.4 集中給気ブロック(２行目)・APS付集中給気ブロック(３行目)チェック
            '8.4.1 使用数設定
            intChargeAirCnt = CInt(strUseValues(Siyou_11.ChargeAir - 1))
            intChargeAirAPSCnt = CInt(strUseValues(Siyou_11.ChargeAirAPS - 1))
            intTotalChargeAirCnt = intChargeAirCnt + intChargeAirAPSCnt
            '8.4.2 選択数判定
            Select Case intTotalChargeAirCnt
                Case 2
                    '集中給気ブロック選択数計が2、且つ、選択オプション(2)が"1","2"の場合、エラー
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2)
                        Case "1", "2"
                            sbCoordinates.Append(CStr(Siyou_11.ChargeAir) & strComma & "0")
                            sbCoordinates.Append(strPipe & CStr(Siyou_11.ChargeAirAPS) & strComma & "0")
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W0620"
                            Exit Function
                        Case Else
                    End Select
                Case 3
                    '集中給気ブロック選択数計が3、且つ、選択オプション(2)が"1","2","3","4"の場合、エラー
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2)
                        Case "1", "2", "3", "4"
                            sbCoordinates.Append(CStr(Siyou_11.ChargeAir) & strComma & "0")
                            sbCoordinates.Append(strPipe & CStr(Siyou_11.ChargeAirAPS) & strComma & "0")
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W0620"
                            Exit Function
                        Case Else
                    End Select
                Case Is >= 4
                    '集中給気ブロック選択数計が4以上の場合、エラー
                    sbCoordinates.Append(CStr(Siyou_11.ChargeAir) & strComma & "0")
                    sbCoordinates.Append(strPipe & CStr(Siyou_11.ChargeAirAPS) & strComma & "0")
                    strMsg = sbCoordinates.ToString
                    strMsgCd = "W0620"
                    Exit Function
            End Select
            '8.5／8.8 レギュレータブロック(４～１３行目)チェック
            For intCI As Integer = 1 To intColR
                For intRI As Integer = Siyou_11.Regulator1 To Siyou_11.Regulator10
                    If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                        '形番"*500B*"チェック
                        If InStr(strKataValues(intRI - 1), "500B") > 0 Then
                            bolRegB = True
                            Exit For
                        End If
                        '"*500B*"選択位置より右列に"*500A*"が選択されていた場合、エラー
                        If InStr(strKataValues(intRI - 1), "500A") > 0 And bolRegB = True Then
                            sbCoordinates.Append("0" & strComma & CStr(intCI))
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W0630"
                            Exit Function
                        End If
                    End If
                Next

                '"*500B*"選択位置より右列に集中給気／APS付集中給気ブロックが選択されていた場合、エラー
                For intRI As Integer = Siyou_11.ChargeAir To Siyou_11.ChargeAirAPS
                    If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                        If bolRegB = True Then
                            sbCoordinates.Append("0" & strComma & CStr(intCI))
                            strMsg = sbCoordinates.ToString
                            strMsgCd = "W0660"
                            Exit Function
                        End If
                    End If
                Next
            Next
            '8.6 シリーズ形番が"MNRB500A","MNRJB500A"の場合
            Select Case strSeriesKata
                Case "MNRB500A", "MNRJB500A"
                    '8.6.1 レギュレータブロックに"*500A*"タイプが一つも選択されていない場合、エラー
                    For intRI As Integer = Siyou_11.Regulator1 To Siyou_11.Regulator10
                        If strUseValues(intRI - 1) > 0 And _
                           InStr(strKataValues(intRI - 1), "500A") > 0 Then
                            bolRegA = True
                        End If
                    Next
                    If bolRegA = False Then
                        strMsgCd = "W0640"
                        Exit Function
                    End If
                Case Else
            End Select
            '8.7 レギュレータブロックに"*500A*"タイプが選択されていた場合
            For intRI As Integer = Siyou_11.Regulator1 To Siyou_11.Regulator10
                If strUseValues(intRI - 1) > 0 And _
                   InStr(strKataValues(intRI - 1), "500A") > 0 Then
                    bolRegA = True
                    Exit For
                End If
            Next
            '集中給気ブロック数(APS付含む)が未選択の場合、エラー
            If bolRegA = True Then
                If intTotalChargeAirCnt = 0 Then
                    strMsgCd = "W0650"
                    Exit Function
                End If
            End If
            '8.9 レギュレータ形番(４～１３行目)・MP付サブベース(１４行目)選択数チェック
            '４～１４行目の選択数カウント
            For intRI As Integer = Siyou_11.Regulator1 To Siyou_11.Subbase
                intSelCnt = intSelCnt + CInt(strUseValues(intRI - 1))
            Next
            Select Case intSelCnt
                Case Is > objKtbnStrc.strcSelection.strOpSymbol(2)
                    '選択数 > 連数 の場合、エラー
                    strMsgCd = "W0670"
                    Exit Function
                Case Is < objKtbnStrc.strcSelection.strOpSymbol(2)
                    '選択数 < 連数の場合、エラー
                    strMsgCd = "W0680"
                    Exit Function
            End Select
            '8.10／8.11 MP付サブベース(１４行目)チェック
            If CInt(strUseValues(Siyou_11.Subbase - 1)) > 0 Then
                For intCI As Integer = intColR To 1 Step -1
                    'MP付サブベースチェック
                    If arySelectInf(Siyou_11.Subbase - 1)(intCI - 1) = "1" Then
                        intSubBaseCol = intCI
                    End If
                    'レギュレータ(３～１３行目)チェック
                    If intSubBaseCol > 0 Then
                        For intRI As Integer = Siyou_11.Regulator1 To Siyou_11.Regulator10
                            If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                                Select Case True
                                    Case InStr(strKataValues(Siyou_11.Subbase - 1), "500A") > 0 And _
                                         InStr(strKataValues(intRI - 1), "500B") > 0
                                        'MP付サブベースが"*500A*"タイプかつ、レギュレータが"*500B*"タイプの場合、エラー
                                        sbCoordinates.Append("0" & strComma & CStr(intSubBaseCol))
                                        strMsg = sbCoordinates.ToString
                                        strMsgCd = "W0690"
                                        Exit Function
                                    Case InStr(strKataValues(Siyou_11.Subbase - 1), "500B") > 0 And _
                                         InStr(strKataValues(intRI - 1), "500A") > 0
                                        'MP付サブベースが"*500B*"タイプかつ、レギュレータが"*500A*"タイプの場合、エラー
                                        sbCoordinates.Append("0" & strComma & CStr(intSubBaseCol))
                                        strMsg = sbCoordinates.ToString
                                        strMsgCd = "W0690"
                                        Exit Function
                                End Select
                            End If
                        Next
                    End If
                Next
            End If
            '8.12 取付レール長さチェック(選択オプション(3)が"D"以外の場合のみ)
            dblRailLen = CDbl(strUseValues(Siyou_11.Rail - 1))
            If objKtbnStrc.strcSelection.strOpSymbol(3) = "D" Then
            Else
                '入力値が範囲外
                If dblRailLen < dblStdNum Or dblRailLen > 600.0 Or _
                   dblRailLen Mod 12.5 <> 0 Then
                    strMsgCd = "W0700"
                    Exit Function
                End If
            End If
            '8.13 MIX構成チェック
            '集中給気ブロック複数選択
            If CInt(strUseValues(Siyou_11.ChargeAir - 1)) > 1 Then
                bolMixFlag = True
            End If
            'APS付集中給気ブロック選択
            If CInt(strUseValues(Siyou_11.ChargeAirAPS - 1)) > 0 Then
                bolMixFlag = True
            End If
            'MP付サブベース選択
            If CInt(strUseValues(Siyou_11.Subbase - 1)) > 0 Then
                bolMixFlag = True
            End If
            '取付レール長さ判定
            If CDbl(strUseValues(Siyou_11.Rail - 1)) <> dblStdNum Then
                bolMixFlag = True
            End If

            '集中給気ブロック設置位置判定
            intChargePos = 0
            For intCI As Integer = 1 To intColR
                If arySelectInf(Siyou_11.ChargeAir - 1)(intCI - 1) = "1" Then
                    intChargePos = intCI
                End If
                If intChargePos > 0 Then
                    For intRI As Integer = Siyou_11.Regulator1 To Siyou_11.Regulator10
                        '左側チェック
                        For intCI2 As Integer = intChargePos - 1 To 1 Step -1
                            If arySelectInf(intRI - 1)(intCI2 - 1) = "1" Then
                                bolRegL = True
                                Exit For
                            End If
                        Next
                        '右側チェック
                        For intCI2 As Integer = intChargePos + 1 To intColR
                            If arySelectInf(intRI - 1)(intCI2 - 1) = "1" Then
                                bolRegR = True
                                Exit For
                            End If
                        Next
                        '集中給気ブロックの両側にレギュレータがある場合、True
                        If bolRegL = True And bolRegR = True Then
                            bolMixFlag = True
                            Exit For
                        End If
                    Next
                End If
                If bolMixFlag = True Then
                    Exit For
                End If
            Next

            'ブランクプラグ使用数判定
            For intRI As Integer = Siyou_11.Plug1 To Siyou_11.Plug3
                If CInt(strUseValues(intRI - 1)) > 0 Then
                    bolMixFlag = True
                    Exit For
                End If
            Next
            'レギュレータブロック
            For intRI As Integer = Siyou_11.Regulator1 To Siyou_11.Regulator10
                '選択種類判定
                If CInt(strUseValues(intRI - 1)) > 0 Then
                    If Trim(strKataValues(intRI - 1)).Length > 0 Then
                        intRegCnt = intRegCnt + 1
                    End If
                    '選択形番判定
                    If InStr(strKataValues(intRI - 1), "U") > 0 Then
                        bolMixFlag = True
                    End If
                End If
            Next
            If intRegCnt > 1 Then
                bolMixFlag = True
            End If

            '上記条件に一つも当てはまらない場合、エラー
            If bolMixFlag = False Then
                strMsgCd = "W0710"
                Exit Function
            End If
            fncInputChk = True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try
    End Function

End Class
