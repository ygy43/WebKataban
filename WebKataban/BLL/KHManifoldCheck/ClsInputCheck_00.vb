Imports WebKataban.ClsCommon

Public Class ClsInputCheck_00
    'message
    Private Structure stcMSG
        Private strDummy As String
        Public Const ACT_LOT As String = "W1750"
        Public Const ACT_LIT As String = "W1760"
        Public Const ACT_LOT2 As String = "W1150"
        Public Const ACT_LIT2 As String = "W1180"
        Public Const NOT_INP As String = "W1030"
        Public Const PLU_SEL As String = "W1740"
        Public Const EXH_TYP As String = "W1770"
        Public Const NOT_PAR As String = "W1780"
    End Structure

    ''' <summary>
    ''' 入力内容ﾁｪｯｸ
    ''' </summary>
    ''' <param name="objKtbnStrc"></param>
    ''' <param name="intRensuu"></param>
    ''' <param name="strMsgCd"></param>
    ''' <param name="StrMsg"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncInputChk(objKtbnStrc As KHKtbnStrc, intRensuu As Long, _
                           ByRef strMsgCd As String, ByRef StrMsg As String) As Boolean
        Dim CST_COMMA As String = CdCst.Sign.Delimiter.Comma
        Dim CST_PIPE As String = CdCst.Sign.Delimiter.Pipe
        Dim sbCollect As New System.Text.StringBuilder
        Dim strCollect() As String
        Dim intCount As Integer      '全体でいくつ選択されているか
        Dim intSelCnt As Integer     '同列にいくつ選択されているか
        Dim intAct As Integer        '連数
        Dim intSelectCount As Integer
        Dim bolChk As Boolean
        Dim strXY As String          'ｴﾗｰ座標

        fncInputChk = True
        Try
            '連数ﾁｪｯｸ
            intCount = 0
            intSelectCount = 0

            'CHANGED BY YGY 20141119
            '位置個数の記録
            Dim strPos As String = objKtbnStrc.strcSelection.strPositionInfo(0)
            For intC As Integer = 1 To strPos.Length
                intSelCnt = 0
                For intR As Integer = 0 To objKtbnStrc.strcSelection.strPositionInfo.Count - 1
                    If CInt(objKtbnStrc.strcSelection.strPositionInfo(intR)(intC - 1).ToString) > 0 Then
                        sbCollect.Append(CStr(intR + 1) & CST_COMMA & intC & CST_PIPE)
                        Select Case objKtbnStrc.strcSelection.strSeriesKataban
                            Case "M4SB0"
                                Select Case intR
                                    Case 1, 2, 3, 4
                                        If objKtbnStrc.strcSelection.strOpSymbol(5) = "C4T31" Or _
                                            objKtbnStrc.strcSelection.strOpSymbol(5) = "D4T30" Or _
                                            objKtbnStrc.strcSelection.strOpSymbol(5) = "C4T50" Then
                                            intCount = intCount + 2
                                        Else
                                            intCount = intCount + 1
                                        End If
                                    Case Else
                                        intCount = intCount + 1
                                End Select
                            Case Else
                                intCount = intCount + 1
                        End Select
                        intSelCnt = intSelCnt + 1
                        intSelectCount = intSelectCount + 1
                    End If
                Next
            Next

            'Dim strPos As String = Nothing
            'For intR As Integer = 0 To objKtbnStrc.strcSelection.strPositionInfo.Count - 1
            '    strPos = objKtbnStrc.strcSelection.strPositionInfo(intR)
            '    intSelCnt = 0
            '    For intC As Integer = 1 To strPos.Length
            '        If CInt(strPos(intC - 1).ToString) > 0 Then
            '            sbCollect.Append(CStr(intR + 1) & CST_COMMA & intC & CST_PIPE)
            '            Select Case objKtbnStrc.strcSelection.strSeriesKataban
            '                Case "M4SB0"
            '                    Select Case intR
            '                        Case 1, 2, 3, 4
            '                            If objKtbnStrc.strcSelection.strOpSymbol(5) = "C4T31" Or _
            '                                objKtbnStrc.strcSelection.strOpSymbol(5) = "D4T30" Or _
            '                                objKtbnStrc.strcSelection.strOpSymbol(5) = "C4T50" Then
            '                                intCount = intCount + 2
            '                            Else
            '                                intCount = intCount + 1
            '                            End If
            '                        Case Else
            '                            intCount = intCount + 1
            '                    End Select
            '                Case Else
            '                    intCount = intCount + 1
            '            End Select
            '            intSelCnt = intSelCnt + 1
            '            intSelectCount = intSelectCount + 1
            '        End If
            '    Next
            'Next

            '連数ﾁｪｯｸ
            'If intRensuu = 0 Then
            '    strMsgCd = "E001"
            '    StrMsg = "連数を確認してください。"
            '    Return False
            'End If
            intAct = objKtbnStrc.strcSelection.strOpSymbol(intRensuu)     '連数
            If objKtbnStrc.strcSelection.strSeriesKataban = "M4SB0" Then
                Dim intSolMax As Integer
                '最大連数を設定
                Select Case objKtbnStrc.strcSelection.strOpSymbol(5)
                    Case "C4T31", "D4T30"
                        intSolMax = 20
                    Case "C4T50"
                        intSolMax = 16
                    Case Else
                        intSolMax = 20
                End Select
                If intSelectCount > intAct Then
                    'message:ｱｸﾁｪｰﾀ数が多すぎます
                    strMsgCd = stcMSG.ACT_LOT
                    Return False
                Else
                    If intSelectCount < intAct Then
                        'message:ｱｸﾁｪｰﾀ数が足りません
                        strMsgCd = stcMSG.ACT_LIT
                        Return False
                    End If
                End If
                If intCount > intSolMax Then
                    'message:ソレノイド点数が多すぎます。
                    strMsgCd = stcMSG.ACT_LOT2
                    Return False
                Else
                    If intCount < intAct Then
                        'message:選択した電磁弁の連数が指定した値より少ないです。
                        strMsgCd = stcMSG.ACT_LIT2
                        Return False
                    End If
                End If
            End If

            If intCount > intAct Then
                If objKtbnStrc.strcSelection.strSeriesKataban <> "M4SB0" Then
                    'message:ｱｸﾁｪｰﾀ数が多すぎます
                    strMsgCd = stcMSG.ACT_LOT
                    Return False
                End If
            End If
            If intCount < intAct Then
                If objKtbnStrc.strcSelection.strSeriesKataban <> "M4SB0" Then
                    'message:ｱｸﾁｪｰﾀ数が足りません
                    strMsgCd = stcMSG.ACT_LIT
                    Return False
                End If
            End If

            If intCount > 1 Then
                strCollect = sbCollect.ToString.Split(CST_PIPE)
                Dim intCol As Integer = 1
                For idx As Integer = 0 To strCollect.Length - 2
                    If intCol = CInt(strCollect(idx).Split(CST_COMMA)(1)) Then
                        intCol = intCol + 1
                    Else
                        'message:設置位置が未入力です。選択してください。
                        strXY = "0" & CST_COMMA & CStr(intCol)
                        strMsgCd = stcMSG.NOT_INP
                        Return False
                    End If
                Next

                For idx As Integer = 0 To strCollect.Length - 3
                    'n列目の選択行数とn+1列目の選択行数を比較する
                    If strCollect(idx).Split(CST_COMMA)(0) = strCollect(idx + 1).Split(CST_COMMA)(0) Then
                        bolChk = False
                    Else
                        bolChk = True
                        Exit For
                    End If
                Next

                'M3QRA,M3QRBでミックス以外のとき電磁弁は1種類でよい
                Select Case objKtbnStrc.strcSelection.strSeriesKataban
                    Case "M3QRA1", "M3QRB1", "MV3QRA1", "MV3QRB1", "M3QB1", "M3QE1", "M3QZ1"
                        If objKtbnStrc.strcSelection.strOpSymbol(1).ToString <> "8" Then
                            bolChk = True
                        End If
                End Select
                If Not bolChk Then
                    'message:電磁弁は2種類以上選択してください
                    strMsgCd = stcMSG.PLU_SEL
                    Return False
                End If
            End If

            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "MV3QRA1", "MV3QRB1"
                    If objKtbnStrc.strcSelection.strOpSymbol(4).ToString.Trim = "" Then
                        For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                            Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(idx).ToString, 6, 1)
                                Case "2"
                                    If CInt(objKtbnStrc.strcSelection.intQuantity(idx).ToString) > 0 Then
                                        strMsgCd = "W9030"
                                        Return False
                                    End If
                            End Select
                        Next
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(5).ToString.Trim = "H" Then
                        For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                            Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(idx).ToString, 6, 1)
                                Case "2"
                                    If CInt(objKtbnStrc.strcSelection.intQuantity(idx).ToString) > 0 Then
                                        strMsgCd = "W9020"
                                        Return False
                                    End If
                            End Select
                        Next
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(7).ToString.Trim = "4" Then
                        For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                            Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(idx).ToString, 6, 1)
                                Case "2"
                                    If CInt(objKtbnStrc.strcSelection.intQuantity(idx).ToString) > 0 Then
                                        strMsgCd = "W9010"
                                        Return False
                                    End If
                            End Select
                        Next
                    End If
                    Dim intKirikae As Integer
                    intKirikae = 0
                    For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                        Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(idx).ToString, 6, 1)
                            Case "1", "2"
                                If CInt(objKtbnStrc.strcSelection.intQuantity(idx).ToString) > 0 Then
                                    If intKirikae = 1 Then
                                        strMsgCd = "W9000"
                                        Return False
                                    End If
                                    intKirikae = 1
                                End If

                        End Select
                    Next
            End Select

            '排気取付方式ﾁｪｯｸ
            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "M4F0", "M4F1", "M4F2"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(8)
                        Case "CL", "IL"
                            For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                                Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(idx).ToString, 4, 1)
                                    Case "2", "3", "4", "5"
                                        If CInt(objKtbnStrc.strcSelection.intQuantity(idx).ToString) > 0 Then
                                            strMsgCd = stcMSG.EXH_TYP
                                            Return False
                                        End If
                                End Select
                            Next
                    End Select
                Case "M4F3"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case ""
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(8)
                                Case "CL", "IL"
                                    For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                                        Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(idx).ToString, 4, 1)
                                            Case "2", "3", "4", "5"
                                                If CInt(objKtbnStrc.strcSelection.intQuantity(idx).ToString) > 0 Then
                                                    strMsgCd = stcMSG.EXH_TYP
                                                    Return False
                                                End If
                                        End Select
                                    Next

                            End Select
                        Case "E"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(10)
                                Case "CL", "IL"

                                    For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                                        Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(idx).ToString, 4, 1)
                                            Case "2", "3", "4", "5"
                                                If CInt(objKtbnStrc.strcSelection.intQuantity(idx).ToString) > 0 Then
                                                    strMsgCd = stcMSG.EXH_TYP
                                                    Return False
                                                End If
                                        End Select
                                    Next

                            End Select
                        Case "X"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(9)
                                Case "CL", "IL"

                                    For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                                        Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(idx).ToString, 4, 1)
                                            Case "2", "3", "4", "5"
                                                If CInt(objKtbnStrc.strcSelection.intQuantity(idx).ToString) > 0 Then
                                                    strMsgCd = stcMSG.EXH_TYP
                                                    Return False
                                                End If
                                        End Select
                                    Next

                            End Select
                    End Select
            End Select

            'ﾌﾟﾗｸﾞ組付ﾁｪｯｸ
            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "M4F1", "M4F2", "M4F3", "M4F4", "M4F5", "M4F6", "M4F7"
                    If objKtbnStrc.strcSelection.strKeyKataban = "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(6).IndexOf("NC") >= 0 Or _
                            objKtbnStrc.strcSelection.strOpSymbol(6).IndexOf("NO") >= 0 Then
                            For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                                Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(idx).ToString, 4, 1)
                                    Case "3", "4", "5"
                                        If CInt(objKtbnStrc.strcSelection.intQuantity(idx).ToString) > 0 Then
                                            strMsgCd = stcMSG.EXH_TYP
                                            Return False
                                        End If
                                End Select
                            Next
                        End If
                    End If

                    If objKtbnStrc.strcSelection.strKeyKataban = "E" Or _
                        objKtbnStrc.strcSelection.strKeyKataban = "X" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(7).IndexOf("NC") >= 0 Or _
                            objKtbnStrc.strcSelection.strOpSymbol(7).IndexOf("NO") >= 0 Then
                            For idx As Integer = 0 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                                Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(idx).ToString, 4, 1)
                                    Case "3", "4", "5"
                                        If CInt(objKtbnStrc.strcSelection.intQuantity(idx).ToString) > 0 Then
                                            strMsgCd = stcMSG.NOT_PAR
                                            Return False
                                        End If
                                End Select
                            Next
                        End If
                    End If
            End Select
            Return True
        Catch ex As Exception
            strMsgCd = "E001"
            StrMsg = ex.Message
            WriteErrorLog(strMsgCd, ex)
        End Try
    End Function
End Class
