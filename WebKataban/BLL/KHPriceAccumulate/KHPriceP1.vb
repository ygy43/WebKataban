'************************************************************************************
'*  ProgramID  ：KHPriceP1
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/10/04   作成者：NII A.Takahashi
'*
'*  概要       ：リフターシリンダ ＬＦＣシリーズ
'*
'************************************************************************************
Module KHPriceP1
    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                       ByRef strOpRefKataban() As String, _
                                       ByRef decOpAmount() As Decimal, _
                                       Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer = 0
        Dim bolC5Flag As Boolean

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'C5チェック
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc)

            'ストローク設定
            intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(1).Trim), _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim))


            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                       "BASE" & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                       intStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
            End If

            'スイッチ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                'リード線長さ加算価格キー
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "T0H", "T0V", "T2H", "T2V", "T3H", "T3V", "T5H", "T5V", "T1H", "T1V", _
                             "T8H", "T8V", "T2YH", "T2YV", "T3YH", "T3YV"
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       CdCst.Sign.Hypen & "SWLW(1)" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case "T2YD"
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       CdCst.Sign.Hypen & "SWLW(2)" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case "T2YDT"
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       CdCst.Sign.Hypen & "SWLW(3)" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case "T2JH", "T2JV"
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       CdCst.Sign.Hypen & "SWLW(4)" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    End Select
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub
End Module
