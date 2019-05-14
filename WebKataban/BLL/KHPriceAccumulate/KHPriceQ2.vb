'************************************************************************************
'*  ProgramID  ：KHPriceQ2
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2008/07/22   作成者：T.Sato
'*                                      更新日：             更新者：
'*
'*  概要       ：オートハンドチェンジャー(ＣＨＣ)
'*
'************************************************************************************
Module KHPriceQ2

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOptionKataban As String = ""

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "R", "H"

                    '基本価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'オプション加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '弁チェック加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "RS", "HS", "S"

                    '基本価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
