'************************************************************************************
'*  ProgramID  ：KHPriceP2
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/10/02   作成者：NII A.Takahashi
'*
'*  概要       ：小形ダイレクトシリンダ 複動・片ロッド形    ＭＤＶシリーズ
'*
'************************************************************************************
Module KHPriceP2

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
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim))


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
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "MDV-L"
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(5).Trim)
                        'リード線長さ加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(5).Trim)
                        End If
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub
End Module
