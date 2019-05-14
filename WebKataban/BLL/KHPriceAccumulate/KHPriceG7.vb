'************************************************************************************
'*  ProgramID  ：KHPriceG7
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：薬液用マニュアルバルブ
'*             ：ＭＭＤ３０２
'*             ：ＭＭＤ４０２
'*             ：ＭＭＤ５０２
'*
'************************************************************************************
Module KHPriceG7

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
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            '流体加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "P" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "*" & _
                                                           Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6, 1) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If


            '取付加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "X" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & _
                                                           CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
