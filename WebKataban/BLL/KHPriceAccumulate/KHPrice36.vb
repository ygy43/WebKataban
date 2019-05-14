'************************************************************************************
'*  ProgramID  ：KHPrice36
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＡＰシリーズ（防爆）
'*             ：ＡＤシリーズ（防爆）
'*             ：ＡＤＫシリーズ（防爆）
'*
'************************************************************************************
Module KHPrice36

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strCountryCd As String = Nothing, _
                                   Optional ByRef strOfficeCd As String = Nothing)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            Select Case True
                Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) = ""
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 2) = "EX"
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6, 2) = "EX"
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 8) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) = "E"
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6, 1) = "E"
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case Else
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 8) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
            End Select
            Select Case True
                Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "H"
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "0"
                Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "J"
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "B"
                Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "K"
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "C"
                Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "L"
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "D"
                Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "M"
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "E"
                Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "N"
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "F"
                Case Else
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1)
            End Select
            If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "5" Or Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "4" Then
                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "3"
            Else
                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1)
            End If
            decOpAmount(UBound(decOpAmount)) = 1

            'コイル別電圧加算価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "AD12E4", "AD22E4", "ADK11E4", "ADK12E4", "AP12E2", "AP12E4"
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "5" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> CdCst.PowerSupply.Const1 And _
                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> CdCst.PowerSupply.Const2 Then

                        '2010/08/27 ADD RM0808112(異電圧対応) START--->
                        If KHKataban.fncVoltageIsStandard(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
                                                        strCountryCd, strOfficeCd) Then
                            '異電圧
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                                Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        'ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        'strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                        '                                                    Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2)
                        'decOpAmount(UBound(decOpAmount)) = 1
                        '2010/08/27 ADD RM0808112(異電圧対応) <--- END
                    End If
                Case "ADK11EX4"
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "5" Or _
                      objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> CdCst.PowerSupply.Const1 And _
                      objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> CdCst.PowerSupply.Const2 Then
                        If KHKataban.fncVoltageIsStandard(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
                                                        strCountryCd, strOfficeCd) Then
                            '異電圧
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                                Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                Case "AD11E4", "AD21E4", "AP11E2", "AP11E4", "AP21E2", "AP21E4", "AP22E2", "AP22E4"
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "5" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> CdCst.PowerSupply.Const1 And _
                       objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> CdCst.PowerSupply.Const2 Then
                        If objKtbnStrc.strcSelection.strSeriesKataban.Trim <> "AP11E4" Then
                            '2010/08/27 ADD RM0808112(異電圧対応) START--->
                            If KHKataban.fncVoltageIsStandard(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
                                                            strCountryCd, strOfficeCd) Then
                                '異電圧
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                                    Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                            'ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            'strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                            '                                                    Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2)
                            'decOpAmount(UBound(decOpAmount)) = 1
                            '2010/08/27 ADD RM0808112(異電圧対応) <--- END
                        Else
                            'コイルオプション加算
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                                Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                Case "AP11EX4"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                        Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 2)
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "AD11EX4", "AD21EX4", "AP11EX2", "AP21EX4", "AP21EX2"
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "5" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> CdCst.PowerSupply.Const1 And _
                            objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> CdCst.PowerSupply.Const2 Then
                        If KHKataban.fncVoltageIsStandard(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                                            strCountryCd, strOfficeCd) Then
                            '異電圧
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                                Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 2)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Else
                        'コイルオプション加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                            Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

            '外部導線引込方式加算価格キー
            Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1)
                Case "L", "M", "N", "P"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "A*4*-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'ボディシール加算価格キー
            Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1)
                Case "H", "J", "K", "L", "M", "N"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション加算価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "AP12E4", "AP12E2", "AD12E4", "AD22E4", "ADK11E4", "ADK12E4", "ADK11EX4"
                Case Else
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(5), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Next
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
