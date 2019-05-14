Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_08

    Public Shared intPosRowCnt As Integer = 16
    Public Shared intColCnt As Integer = 25

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
            WriteErrorLog("E001", ex)
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
    Public Shared Function fncInpCheck2(objKtbnStrc As KHKtbnStrc, ByRef dblStdNum As Double, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean

        Dim sbCoordinates As New System.Text.StringBuilder
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

            ''エンドプレート複数不可チェック
            If CInt(strUseValues(Siyou_08.EndP1 - 1)) > 1 Then
                sbCoordinates.Append(CStr(Siyou_08.EndP1) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W2310"
                Exit Function
            End If
            If CInt(strUseValues(Siyou_08.EndP2 - 1)) > 1 Then
                sbCoordinates.Append(CStr(Siyou_08.EndP2) & strComma & "0")
                strMsg = sbCoordinates.ToString
                strMsgCd = "W2320"
                Exit Function
            End If

            '' 品名リストコントロール 数値テキスト入力値チェック
            If Not SiyouBLL.fncOtherKataCheck(objKtbnStrc, Siyou_08.Silencer1, Siyou_08.Inspect1, _
                                     Siyou_08.Rail, strMsgCd) Then
                Exit Function
            End If

            ''取付レール長さ設定値チェック
            If strKataValues(Siyou_08.Rail - 1).ToString.Length <= 0 Then strKataValues(Siyou_08.Rail - 1) = 0
            If Not SiyouBLL.fncRailchk(strKataValues(Siyou_08.Rail - 1), strUseValues(Siyou_08.Rail - 1), dblStdNum, strMsgCd) Then
                strMsg = Siyou_08.Rail & ",0"
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
        Dim intMixCon(4) As Integer
        Dim bolMixSwtch(6) As Boolean
        Dim intMixConCnt As Integer = 0
        Dim intMixSwtchCnt As Integer = 0
        Dim intElectSeq As Integer
        Dim strKataban As String
        Dim strKataSub As String

        fncInpCheck1 = False
        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            ''左側エンドプレートチェック
            If strKataValues(0) = "" Then
                Exit Function
            End If

            ''接続位置チェック
            For intCI As Integer = intColCnt To 1 Step -1
                bolPosChkFlg = False

                For intRI As Integer = 1 To intPosRowCnt
                    If bolRightChkFlg = False Then
                        If arySelectInf(intRI - 1)(intCI - 1) = "1" Then
                            If intRI = Siyou_08.EndP2 Then
                                bolRightChkFlg = True
                                bolPosChkFlg = True
                            Else
                                sbCoordinates.Append("0" & strComma & CStr(intCI))
                                strMsg = sbCoordinates.ToString
                                strMsgCd = "W2330"
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

            ''********** 電磁弁連数チェック ******************************
            For intI As Integer = 0 To UBound(intMixCon)
                intMixCon(intI) = 0
            Next
            For intI As Integer = 0 To UBound(bolMixSwtch)
                bolMixSwtch(intI) = False
            Next
            intElectSeq = 0
            '３～１２行目チェック
            For intRI As Integer = Siyou_08.ElType1 - 1 To Siyou_08.ElType10 - 1
                '対象行が１つ以上選択されている場合
                If CInt(strUseValues(intRI)) > 0 Then
                    strKataban = Trim(strKataValues(intRI))
                    '形番要素が選択されている場合
                    If Len(strKataban) > 0 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).ToString = "8" Then
                            strKataSub = strKataban.Substring(3, 1)
                            Select Case strKataSub
                                Case "2", "3", "4"
                                    intMixCon(CInt(strKataSub)) = intMixCon(CInt(strKataSub)) + CInt(strUseValues(intRI))
                            End Select
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(4).ToString = "8" Then
                            strKataSub = strKataban.Substring(4, 1)
                            Select Case strKataSub
                                Case "2", "6"
                                    bolMixSwtch(CInt(strKataSub)) = True
                            End Select
                        End If

                        '電磁弁連数値セット
                        intElectSeq = intElectSeq + Int(strUseValues(intRI))

                    End If
                End If
            Next

            'エラーチェック
            If intElectSeq > CInt(objKtbnStrc.strcSelection.strOpSymbol(1)) Then
                strMsgCd = "W1170"
                Exit Function
            End If

            If intElectSeq < CInt(objKtbnStrc.strcSelection.strOpSymbol(1)) Then
                strMsgCd = "W1180"
                Exit Function
            End If

            If objKtbnStrc.strcSelection.strOpSymbol(3).ToString = "8" Then
                For intI As Integer = 0 To UBound(intMixCon)
                    If intMixCon(intI) > 0 Then
                        intMixConCnt = intMixConCnt + 1
                    End If
                Next
                If intMixConCnt <= 1 Then
                    strMsgCd = "W1740"
                    Exit Function
                End If
            End If

            If objKtbnStrc.strcSelection.strOpSymbol(4).ToString = "8" Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).ToString
                    Case "2", "3", "8"
                        If bolMixSwtch(2) = False Or bolMixSwtch(6) = False Then
                            strMsgCd = "W2340"
                            Exit Function
                        End If
                End Select
            End If

            ''********** 給排気チェック ******************************
            For intRI As Integer = Siyou_08.EndP1 - 1 To Siyou_08.EndP2 - 1
                strKataban = Trim(strKataValues(intRI))
                If CInt(strUseValues(intRI)) > 0 Then
                    If strKataban.Substring(0, 1) = "P" Then
                        bolExhaustFlg1 = True
                    End If
                    If strKataban.Substring(0, 1) = "R" Then
                        bolExhaustFlg2 = True
                    End If
                End If
            Next
            For intRI As Integer = Siyou_08.Supply1 - 1 To Siyou_08.Exhaust2 - 1
                strKataban = Trim(strKataValues(intRI))
                If CInt(strUseValues(intRI)) > 0 Then
                    If strKataban.Substring(0, 1) = "P" Then
                        bolExhaustFlg1 = True
                    End If
                    If strKataban.Substring(0, 1) = "R" Then
                        bolExhaustFlg2 = True
                    End If
                End If
            Next

            If bolExhaustFlg1 = False Then
                strMsgCd = "W2350"
                Exit Function
            End If
            If bolExhaustFlg2 = False Then
                strMsgCd = "W2360"
                Exit Function
            End If

            fncInpCheck1 = True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try

    End Function
End Class
