Module KHCylinderSTSCheck

    '********************************************************************************************
    '*【関数名】
    '*  fncCheckSelectOption
    '*【処理】
    '*  シリンダチェック
    '*【概要】
    '*  シリンダＳＴＳ－Ｌシリーズをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2007/05/16      更新者：NII A.Takahashi
    '*  ・T2W/T3Wスイッチ追加に伴い、ストロークチェックロジックを修正
    '********************************************************************************************
    Public Function fncCheckSelectOption(ByVal objKtbnStrc As KHKtbnStrc, _
                                         ByRef intKtbnStrcSeqNo As Integer, _
                                         ByRef strOptionSymbol As String, _
                                         ByRef strMessageCd As String) As Boolean

        Dim objPrice As New KHUnitPrice

        Try

            fncCheckSelectOption = True

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "STS-B", "STS-M", "STL-B", "STL-M"
                    '*-----<< Ⅰ．最小ストロークチェック >>-----*
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                        Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                             "T3V", "T5H", "T5V", "T2YH", "T2YV", _
                             "T3YH", "T3YV", "T2YFH", "T2YFV", "T3YFH", _
                             "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV", _
                             "T2JH", "T2JV", "T2YD", "T2YDT", "T1H", "T1V", "T8H", "T8V", _
                             "T2WH", "T2WV", "T3WH", "T3WV"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 5 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncCheckSelectOption = False
                                Exit Try
                            End If
                        Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 5 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncCheckSelectOption = False
                                Exit Try
                            End If
                        Case "ET0H", "ET0V"
                            ' スイッチ個数で判定
                            Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(11).Trim)
                                Case "1" ' スイッチ１個（"R"/"H"）
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W0200"
                                        fncCheckSelectOption = False
                                        Exit Try
                                    End If
                                Case Else ' スイッチ２個・３個（"D"/"T"）
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 15 Then
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W0200"
                                        fncCheckSelectOption = False
                                        Exit Try
                                    End If
                            End Select
                        Case "H0", "H0Y"
                            ' スイッチ個数で判定
                            Select Case KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(11).Trim)
                                Case "1" ' スイッチ１個（"R"/"H"）
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 20 Then
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W0200"
                                        fncCheckSelectOption = False
                                        Exit Try
                                    End If
                                Case Else ' スイッチ２個・３個（"D"/"T"）
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 50 Then
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W0200"
                                        fncCheckSelectOption = False
                                        Exit Try
                                    End If
                            End Select
                    End Select

                    '*-----<< Ⅱ．最大ストロークチェック >>-----*
                    Select Case Left(objKtbnStrc.strcSelection.strSeriesKataban, 3)
                        Case "STS"
                            ' バリエーション毎のチェック
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                Case "C", "PC", "CT", "CT1", "CT2", _
                                     "CO", "CG", "CG1", "CG2", "CG3", _
                                     "CG4", "CTG1", "CT1G1", "CT2G1", "PCT2", _
                                     "PCO", "PCG", "PCG1", "PCG4", "PCT2G1"
                                    ' 口径毎のチェック
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "25", "32", "40", "50", "63"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 50 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                        Case "80"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 100 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                    End Select
                                Case "Q", "PQ"
                                    ' 口径毎のチェック
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "25", "32", "40", "50", "63"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 50 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                        Case "80"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 100 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                    End Select
                                Case "F"
                                    ' 口径毎のチェック
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "8", "12", "16", "20", "25", _
                                             "32", "40", "50", "63"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 50 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                        Case "80"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 100 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                    End Select
                                Case "V1", "V2", "V1S", "V2S", "PV1", _
                                     "PV2", "PV1S", "PV2S", "PV1O", "PV2O", _
                                     "PV1SO", "PV2SO"
                                    ' 口径毎のチェック
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "20", "25", "32", "40", "50", "63"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 50 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                    End Select
                                Case Else
                                    ' 口径毎のチェック
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "8", "12", "16", "20", "25", _
                                             "32", "40", "50", "63"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 50 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                        Case "80", "100"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 100 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                    End Select
                            End Select
                        Case "STL"
                            ' バリエーション毎のチェック
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                Case "C", "PC", "CT", "CT1", "CT2", _
                                     "CG", "CG1", "CG2", "CG3", _
                                     "CG4", "CTG1", "CT1G1", "CT2G1", "PCT2", _
                                     "PCG", "PCG1", "PCG4", "PCT2G1"
                                    ' 口径毎のチェック
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "25", "32", "40", "50", "63"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 400 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                        Case "80"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 375 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                    End Select
                                Case "Q", "PQ"
                                    ' 口径毎のチェック
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "20", "25", "32"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 400 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                        Case "40", "50", "63"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 375 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                        Case "80"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 350 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                    End Select
                                Case "F"
                                    ' 口径毎のチェック
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "8", "12", "16", "20", "25", _
                                             "32", "40", "50", "63", "80"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 200 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                    End Select
                                Case "V1", "V2", "V1S", "V2S", "PV1", _
                                     "PV2", "PV1S", "PV2S", "PV1O", "PV2O", _
                                     "PV1SO", "PV2SO"
                                    ' 口径毎のチェック
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "20", "25", "32", "40", "50", _
                                             "63"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 100 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                    End Select
                                Case "O", "PO", "CO", "PCO"
                                    ' 口径毎のチェック
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "8", "12", "16"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 150 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                        Case "20", "25", "32", "40", "50", "63", "80"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 300 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                        Case "100"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 200 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                    End Select
                                Case Else
                                    ' 口径毎のチェック
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "8", "12", "16"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 200 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                        Case "20", "25", "32", "40", "50", _
                                             "63", "80"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 400 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                        Case "100"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 200 Then
                                                intKtbnStrcSeqNo = 5
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                                Exit Try
                                            End If
                                    End Select
                            End Select
                    End Select

                    '*-----<< Ⅲ．中間ストロークチェック >>-----*
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "", "Q", "V1", "V2", "V1S", _
                             "V2S", "T", "T1", "T1L", "T2", _
                             "O", "G", "G1", "G2", "G3", _
                             "G4", "F", "TG1", "T1G1", "T1LG1", _
                             "T2G1"
                            If Right(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 1) = "0" Or _
                                               Right(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 1) = "5" Then
                            Else
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0510"
                                fncCheckSelectOption = False
                                Exit Try
                            End If
                        Case Else
                            Select Case Left(objKtbnStrc.strcSelection.strSeriesKataban, 3)
                                Case "STS"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                        Case "P", "PT2", "PO", "PG", "PG1", _
                                             "PG4"
                                            ' 口径毎のチェック
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "8", "12", "16"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "10", "20", "30", "40", "50"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                                Case "20", "25", "32", "40", "50", _
                                                     "63"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "25", "50"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                                Case "80", "100"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "25", "50", "75", "100"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                            End Select
                                        Case "C", "PC", "CT", "CT1", "CT2", _
                                             "CO", "CG", "CG1", "CG2", "CG3", _
                                             "CG4", "CTG1", "CT1G1", "CT2G1", "PCT2", _
                                             "PCO", "PCG", "PCG1", "PCG4", "PCT2G1"
                                            ' 口径毎のチェック
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "25", "32", "40", "50", "63"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "25", "50"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                                Case "80"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "25", "50", "75", "100"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                            End Select
                                        Case "PQ"
                                            ' 口径毎のチェック
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "20", "25", "32", "40", "50", _
                                                     "63"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "25", "50"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                                Case "80"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "25", "50", "75", "100"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                            End Select
                                        Case "PV1", "PV2", "PV1S", "PV2S", "PV1O", _
                                             "PV2O", "PV1SO", "PV2SO"
                                            ' 口径毎のチェック
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "20", "25", "32", "40", "50", _
                                                     "63"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "25", "50"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                            End Select
                                    End Select
                                Case Else
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                        Case "P", "PT2", "PO", "PG", "PG1", _
                                             "PG4"
                                            ' 口径毎のチェック
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "8", "12", "16"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "50", "75", "100", "125", "150", _
                                                             "175", "200"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                                Case "20", "25", "32", "40", "50", _
                                                     "63"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "50", "75", "100", "125", "150", _
                                                             "175", "200", "225", "250", "275", _
                                                             "300", "325", "350", "375", "400"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                                Case "80"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "75", "100", "125", "150", "175", _
                                                             "200", "225", "250", "275", "300", _
                                                             "325", "350", "375", "400"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                                Case "100"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "75", "100", "125", "150", "175", _
                                                             "200"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                            End Select
                                        Case "C", "PC", "CT", "CT1", "CT2", _
                                             "CO", "CG", "CG1", "CG2", "CG3", _
                                             "CG4", "CTG1", "CT1G1", "CT2G1", "PCT2", _
                                             "PCO", "PCG", "PCG1", "PCG4", "PCT2G1"
                                            ' 口径毎のチェック
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "25", "32", "40", "50", "63"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "50", "75", "100", "125", "150", _
                                                             "175", "200", "225", "250", "275", _
                                                             "300", "325", "350", "375", "400"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                                Case "80"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "75", "100", "125", "150", "175", _
                                                             "200", "225", "250", "275", "300", _
                                                             "325", "350", "375"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                            End Select
                                        Case "PQ"
                                            ' 口径毎のチェック
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "20", "25", "32"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "50", "75", "100", "125", "150", _
                                                             "175", "200", "225", "250", "275", _
                                                             "300", "325", "350", "375", "400"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                                Case "40", "50", "63"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "50", "75", "100", "125", "150", _
                                                             "175", "200", "225", "250", "275", _
                                                             "300", "325", "350", "375"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                                Case "80"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "75", "100", "125", "150", "175", _
                                                             "200", "225", "250", "275", "300", _
                                                             "325", "350"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                            End Select
                                        Case "PV1", "PV2", "PV1S", "PV2S", "PV1O", _
                                             "PV2O", "PV1SO", "PV2SO"
                                            ' 口径毎のチェック
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "20", "25", "32", "40", "50", _
                                                     "63"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "50", "75", "100"
                                                        Case Else
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0590"
                                                            fncCheckSelectOption = False
                                                            Exit Try
                                                    End Select
                                            End Select
                                    End Select
                            End Select
                    End Select
            End Select

        Catch ex As Exception

            Throw ex

        Finally

            objPrice = Nothing

        End Try

    End Function

End Module
