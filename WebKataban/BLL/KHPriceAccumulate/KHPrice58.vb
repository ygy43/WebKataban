'************************************************************************************
'*  ProgramID  ：KHPrice58
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/05   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：省配線バルブ　Ｎ３Ｓ０／Ｎ４Ｓ０
'*
'************************************************************************************
Module KHPrice58

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
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            'A・Bポートフィルタ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Or objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "11" Then
                    strOpRefKataban(UBound(strOpRefKataban)) = "N3S0" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    strOpRefKataban(UBound(strOpRefKataban)) = "N4S0" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            '手動装置加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "N4S0" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                '切換位置区分が"1","11"の時,数量は１
                If Mid(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1, 1) = 1 Then
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    decOpAmount(UBound(decOpAmount)) = 2
                End If
            End If

            '配線方式（個別コネクタ）加算価格キー
            '配線方式が"C","C0","C1","C2"の時
            If Mid(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 1, 1) = "C" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "N4S0" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        End Try


    End Sub

End Module
