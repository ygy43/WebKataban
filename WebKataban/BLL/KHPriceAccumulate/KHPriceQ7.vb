'************************************************************************************
'*  ProgramID  ：KHPriceQ7
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2009/04/06   作成者：T.Yagyu
'*                                      更新日：             更新者：
'*
'*  概要       ：FSLシリーズ
'*
'************************************************************************************
Module KHPriceQ7

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
            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
            objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
            objKtbnStrc.strcSelection.strOpSymbol(2)
            decOpAmount(UBound(decOpAmount)) = 1

            'オプション加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & CdCst.Sign.Hypen & _
                objKtbnStrc.strcSelection.strOpSymbol(3)
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module

