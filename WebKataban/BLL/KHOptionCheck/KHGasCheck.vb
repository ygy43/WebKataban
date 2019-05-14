Module KHGasCheck

    '********************************************************************************************
    '*【関数名】
    '*  fncCheckSelectOption
    '*【処理】
    '*  ガス燃焼システムチェック
    '*【概要】
    '*  ガス燃焼システムをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Public Function fncCheckSelectOption(ByVal objKtbnStrc As KHKtbnStrc, _
                                         ByRef intKtbnStrcSeqNo As Integer, _
                                         ByRef strOptionSymbol As String, _
                                         ByRef strMessageCd As String) As Boolean

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncCheckSelectOption = True

            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "VNA"
                    If objKtbnStrc.strcSelection.strKeyKataban = "" Then
                        Dim bolOptionH As Boolean = False
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "H"
                                    bolOptionH = True
                            End Select
                        Next

                        '接続口径判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            Case "50", "65"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "AC24V", "DC12V"
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W8020"
                                        fncCheckSelectOption = False
                                End Select
                            Case "32", "40"
                                If bolOptionH = True Then
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "AC24V", "DC12V"
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                End If
                        End Select
                    Else
                        '接続口径判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            Case "32", "40"
                                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "H" Then
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "AC24V", "DC12V"
                                            intKtbnStrcSeqNo = 4
                                            strMessageCd = "W8020"
                                            fncCheckSelectOption = False
                                    End Select
                                End If
                        End Select
                    End If
                Case "VLA"
                    '接続口径判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "50", "65"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "AC24V", "DC12V"
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W8020"
                                    fncCheckSelectOption = False
                            End Select
                    End Select
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

End Module
