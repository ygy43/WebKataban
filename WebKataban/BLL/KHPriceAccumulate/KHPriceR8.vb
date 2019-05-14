'************************************************************************************
'*  ProgramID  ：KHPriceR8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2012/04/25   作成者：Y.Tachi
'*                                      更新日：             更新者：
'*
'*  概要       ：精密レギュレータ             ＲＰＥ１０００
'*
'************************************************************************************
Module KHPriceR8

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim intStroke As Integer = 0
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2)
            decOpAmount(UBound(decOpAmount)) = 1

            'オプション加算価格キー
            '2016/2/18 斉藤修正
            'If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) <> " " Then
            If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) <> "" Then
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & "-" & strOpArray(intLoopCnt).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Next
            End If

            '二次電池
            If objKtbnStrc.strcSelection.strKeyKataban.ToString = "4" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4)
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module
