'************************************************************************************
'*  ProgramID  ：KHPriceA8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：インデックスマン　ＲＧＩＢ
'*
'************************************************************************************
Module KHPriceA8

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            'オプション加算価格キー
            For intLoopCnt = 8 To 9
                Select Case objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                    Case "C", "F", "P"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                                   Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "H"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                                   Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(12)
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next


        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
