'************************************************************************************
'*  ProgramID  ：KHPriceC0
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：カートリッジシリンダ　ＣＡＴ
'*
'************************************************************************************
Module KHPriceC0

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
            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            '支持形式加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "N"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
