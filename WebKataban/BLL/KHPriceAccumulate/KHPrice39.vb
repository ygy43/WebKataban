'************************************************************************************
'*  ProgramID  ：KHPrice39
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＨＶＬ（ファイン）
'*
'************************************************************************************
Module KHPrice39

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                       Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2)
            decOpAmount(UBound(decOpAmount)) = 1

            'Aポート加算価格
            If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) <> "6" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "" Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            'Bポート加算価格
            If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) <> "6" And _
               Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
