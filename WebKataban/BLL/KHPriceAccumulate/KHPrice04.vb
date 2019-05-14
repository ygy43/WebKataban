'************************************************************************************
'*  ProgramID  ：KHPrice04
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＡＢ２の単価計算を行う
'*
'************************************************************************************
Module KHPrice04

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strCountryCd As String = Nothing, _
                                   Optional ByRef strOfficeCd As String = Nothing)



        Dim strStdVoltageFlag As String = CdCst.VoltageDiv.Standard
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "BB41", "BG41", "GBB41", "GBG41"
                    '価格キー設定
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                Case Else
                    '価格キー設定
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'オプション価格
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                '価格キー設定
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_04" & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    '2010/08/25 MOD RM0808112(海外異電圧削除対応) START --->
                    '電圧加算
                    strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
                                                                   strCountryCd, strOfficeCd)
                    'strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                    '                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim)
                    '2010/08/25 MOD RM0808112(海外異電圧削除対応) <--- END

                    Select Case strStdVoltageFlag
                        Case CdCst.VoltageDiv.Standard
                        Case CdCst.VoltageDiv.Options
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_051"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case CdCst.VoltageDiv.Other
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_052"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
            End Select
        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
