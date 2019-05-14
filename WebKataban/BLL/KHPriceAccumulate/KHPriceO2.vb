'************************************************************************************
'*  ProgramID  ：KHPriceO2
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/04/18   作成者：NII A.Tatakashi
'*                                      更新日：             更新者：
'*
'*  概要       ：エアハイドロブースタ ＡＨＢシリーズ
'*
'************************************************************************************
Module KHPriceO2

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)
        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)

            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
            decOpAmount(UBound(decOpAmount)) = 1

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
