'************************************************************************************
'*  ProgramID  ：KHPriceQ3
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2008/08/11   作成者：M.KOJIMA
'*                                      更新日：             更新者：
'*
'*  概要       ：ロータリジョイントRJFシリーズ
'*
'************************************************************************************
Module KHPriceQ3

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOptionKataban As String = ""

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            If Len(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3)
                decOpAmount(UBound(decOpAmount)) = 1
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module
