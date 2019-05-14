Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_12

    Public Shared intPosRowCnt As Integer = 10
    Public Shared intColCnt As Integer = 10

    Public Shared Function fncInputChk(objKtbnStrc As KHKtbnStrc, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        fncInputChk = False
        Try
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

    '*   fncInpCheck
    '*【処理】

    '*   入力チェック
    '*【引数】

    '*   strKataValues  : 形番の選択値配列          strUseValues    : 使用数の入力値配列
    '*   arySelectInf   : 設置位置の選択値配列      arySelectInf    ：設置位置の選択値配列(中間行)
    '********************************************************************************************
    Public Shared Function fncInpCheck1(objKtbnStrc As KHKtbnStrc, _
                                       ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        fncInpCheck1 = False

        Dim intMaxSeq As Integer = 0
        Dim intCnt As Integer = 0
        Dim strSeriesKata As String = objKtbnStrc.strcSelection.strSeriesKataban

        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            '連数設定
            Select Case strSeriesKata
                Case "VSJM", "VSXM", "VSZM", "VSNM", "VSNM"
                    intMaxSeq = Int(objKtbnStrc.strcSelection.strOpSymbol(8))
                Case "VSJPM"
                    intMaxSeq = Int(objKtbnStrc.strcSelection.strOpSymbol(7))
                Case "VSXPM"
                    intMaxSeq = Int(objKtbnStrc.strcSelection.strOpSymbol(6))
                Case "VSKM"
                    intMaxSeq = Int(objKtbnStrc.strcSelection.strOpSymbol(9))
                Case "VSZPM", "VSNPM"
                    intMaxSeq = Int(objKtbnStrc.strcSelection.strOpSymbol(5))
            End Select

            Select Case strSeriesKata
                Case "VSNM", "VSNPM"
                    intColCnt = 10
                Case Else
                    intColCnt = 12
            End Select

            '設置位置重複チェック
            For intCI As Integer = 0 To intColCnt - 1
                intCnt = 0
                For intRI As Integer = 0 To arySelectInf.Count - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        intCnt = intCnt + 1
                    End If
                Next

                If intCnt = 0 And intCI < intMaxSeq Then
                    strMsgCd = "W1020"
                    strMsg = CStr(intCI + 1)
                    Exit Try
                ElseIf intCnt > 0 And intCI >= intMaxSeq Then
                    strMsgCd = "W2500"
                    Exit Try
                End If
            Next

            intCnt = 0
            For intRI As Integer = Siyou_12.Vaccum1 - 1 To Siyou_12.Vaccum8 - 1
                If CInt(strUseValues(intRI)) > 0 Then
                    intCnt = intCnt + CInt(strUseValues(intRI))
                End If
            Next

            '真空ユニット、真空切替エジェクタ選択チェック
            If intCnt = 0 And strSeriesKata = "VSKM" Then
                strMsgCd = "W2530"
                Exit Try
            End If

            '形番重複チェック
            For inti As Integer = 0 To strKataValues.Count - 1
                If strKataValues(inti).ToString.Length > 0 Then
                    For intj As Integer = inti + 1 To strKataValues.Count - 1
                        If strKataValues(inti) = strKataValues(intj) Then
                            strMsgCd = "W1330"
                            Exit Try
                        End If
                    Next
                End If
            Next

            Dim intIdx As Integer = 0
            Select Case strSeriesKata
                Case "VSJM"
                    intIdx = 10
                Case "VSXM", "VSZM", "VSJPM"
                    intIdx = 9
                Case "VSXPM"
                    intIdx = 7
            End Select

            Select Case strSeriesKata
                Case "VSJM", "VSXM", "VSZM"
                    'ミックス構成チェック
                    If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                        If Not fncMixCompCheck(objKtbnStrc, "3,4") Then
                            strMsgCd = "W2540"
                            Exit Try
                        End If
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(3) = "Z" Then
                        If Not fncMixCompCheck(objKtbnStrc, "5") Then
                            strMsgCd = "W2560"
                            Exit Try
                        End If
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(4) = "CX" Then
                        If Not fncMixCompCheck(objKtbnStrc, "7") Then
                            strMsgCd = "W2550"
                            Exit Try
                        End If
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(intIdx) = "Z" Then
                        If Not fncMixCompCheck(objKtbnStrc, "9") Then
                            strMsgCd = "W2570"
                            Exit Try
                        End If
                    End If
                Case "VSJPM", "VSXPM"
                    'ミックス構成チェック
                    If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                        If Not fncMixCompCheck(objKtbnStrc, "3") Then
                            strMsgCd = "W2560"
                            Exit Try
                        End If
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(2) = "CX" Then
                        If Not fncMixCompCheck(objKtbnStrc, "4") Then
                            strMsgCd = "W2550"
                            Exit Try
                        End If
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(intIdx) = "Z" Then
                        If Not fncMixCompCheck(objKtbnStrc, "6") Then
                            strMsgCd = "W2570"
                            Exit Try
                        End If
                    End If
                Case "VSKM"
                    For intRI As Integer = Siyou_12.Vaccum1 - 1 To Siyou_12.Vaccum8 - 1
                        If Int(strUseValues(intRI)) > 0 Then
                            'ミックス構成チェック
                            If objKtbnStrc.strcSelection.strOpSymbol(1) = "Z" Then
                                If Not fncMixCompCheck(objKtbnStrc, "3,4") Then
                                    strMsgCd = "W2540"
                                    Exit Try
                                End If
                            End If

                            If objKtbnStrc.strcSelection.strOpSymbol(3) = "Z" Then
                                If Not fncMixCompCheck(objKtbnStrc, "5") Then
                                    strMsgCd = "W2560"
                                    Exit Try
                                End If
                            End If

                            If objKtbnStrc.strcSelection.strOpSymbol(4) = "CX" Then
                                If Not fncMixCompCheck(objKtbnStrc, "7") Then
                                    If CInt(strUseValues(Siyou_12.Mask1 - 1)) = 0 And _
                                       CInt(strUseValues(Siyou_12.Mask2 - 1)) = 0 Then
                                        strMsgCd = "W2550"
                                        Exit Try
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case "VSZPM", "VSNPM"
                    'ミックス構成チェック
                    If objKtbnStrc.strcSelection.strOpSymbol(1) = "CX" Then
                        If Not fncMixCompCheck(objKtbnStrc, "3") Then
                            strMsgCd = "W2550"
                            Exit Try
                        End If
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(6) = "Z" Then
                        If Not fncMixCompCheck(objKtbnStrc, "5") Then
                            strMsgCd = "W2570"
                            Exit Try
                        End If
                    End If
            End Select

            fncInpCheck1 = True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】
    '*   fncMixCompCheck
    '*【処理】
    '*   ミックス構成チェック
    '********************************************************************************************
    Public Shared Function fncMixCompCheck(objKtbnStrc As KHKtbnStrc, ByVal strOptIdx As String) As Boolean
        Dim strIdx() As String
        Dim strKey As String = ""
        Dim hshtSelOpt As New ArrayList

        Try
            strIdx = strOptIdx.Split(strComma)

            For intRI As Integer = Siyou_12.Vaccum1 - 1 To Siyou_12.Vaccum8 - 1
                If objKtbnStrc.strcSelection.strOptionKataban(intRI).ToString.Length <= 0 Then Continue For
                Dim str() As String = objKtbnStrc.strcSelection.strOptionKataban(intRI).ToString.Split(strComma)
                If str.Length <= 0 Then Continue For
                For intI As Integer = 0 To strIdx.Length - 1
                    If intI > 0 Then
                        strKey = strKey & strComma & str(Int(strIdx(intI)) - 1)
                    Else
                        strKey = str(Int(strIdx(intI)) - 1)
                    End If
                Next
                If Int(objKtbnStrc.strcSelection.intQuantity(intRI)) > 0 Then
                    If Len(strKey) = 0 Then
                        strKey = " "
                    End If
                    If hshtSelOpt.Contains(strKey) Then
                        Continue For
                    Else
                        hshtSelOpt.Add(strKey)
                    End If
                End If
            Next

            If hshtSelOpt.Count > 1 Then
                fncMixCompCheck = True
            Else
                fncMixCompCheck = False
            End If
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

    End Function

End Class
