'************************************************************************************
'*  ProgramID  ：KHPriceM0
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：インデックスマン　ＲＧ＊＊／ＰＣ＊Ｓ
'*             ：（ＨＯ減速機取付用インデックスマン本体形番用）
'*
'************************************************************************************
Module KHPriceM0

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            If Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 1) >= "0" And _
               Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 1) <= "9" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                           Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & CdCst.Sign.Hypen & "W" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "FC"
                decOpAmount(UBound(decOpAmount)) = 1
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                           Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & CdCst.Sign.Hypen & "W" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "AL"
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'ＨＯ減速機取付用加算価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "HO"
            decOpAmount(UBound(decOpAmount)) = 1

            Select Case Left(Trim(objKtbnStrc.strcSelection.strSeriesKataban.Trim), 4)
                Case "RGCS", "RGIL", "RGIS", "RGOL", "RGOS", "PCIS", "PCOS"
                    If objKtbnStrc.strcSelection.strOpSymbol(12).Trim <> "" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            Case "F", "A", "S", "B"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                                           Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "TSF" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "X", "C", "Y", "D"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                                           Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "TGX" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
