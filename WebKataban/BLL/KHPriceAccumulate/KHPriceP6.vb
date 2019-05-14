'************************************************************************************
'*  ProgramID  ：KHPriceP6
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2008/06/09   作成者：M.Kojima
'*
'*  概要       ：落下防止付扁平シリンダ　UFCDシリーズ
'*
'************************************************************************************
Module KHPriceP6
    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer = 0
        Dim strBoreSize As String           '口径
        Dim strStroke As String             'ストローク

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            strBoreSize = objKtbnStrc.strcSelection.strOpSymbol(1).Trim
            strStroke = objKtbnStrc.strcSelection.strOpSymbol(3).Trim

            '可変ストローク設定　
            intStroke = _
                KHKataban.fncGetStrokeSize(objKtbnStrc, _
                    CInt(strBoreSize), CInt(strStroke))

            '基本価格キーの設定
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = _
                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                intStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1

            'マグネット内臓(L)加算
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = _
                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                "L"
            decOpAmount(UBound(decOpAmount)) = 1

            'スイッチ加算価格キー

            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim.Length <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                    objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                    objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                'リード線長さ加算価格キー
                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim.Length <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
