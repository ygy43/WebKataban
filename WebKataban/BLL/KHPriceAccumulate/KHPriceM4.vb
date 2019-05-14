'************************************************************************************
'*  ProgramID  ：KHPriceM4
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/25   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ジャスフィットバルブ
'*             ：ＦＶＢ（クリーンブロー用）
'*
'************************************************************************************
Module KHPriceM4

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strStdVoltageFlag As String

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            'クリーン仕様加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "P90" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'コイルオプション加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                Case "2G", "4A"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'その他オプション加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                'クリーン仕様時
                If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "P90" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            '電圧加算価格キー
            strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                           objKtbnStrc.strcSelection.strOpSymbol(10).Trim)
            Select Case strStdVoltageFlag
                Case CdCst.VoltageDiv.Standard
                Case CdCst.VoltageDiv.Options
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "OPT" & CdCst.Sign.Hypen & _
                                                               Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2)
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CdCst.VoltageDiv.Other
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "OTH" & CdCst.Sign.Hypen & _
                                                               Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2)
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
