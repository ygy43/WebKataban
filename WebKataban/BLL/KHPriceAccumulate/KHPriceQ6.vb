'************************************************************************************
'*  ProgramID  ：KHPriceQ6
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2009/03/05   作成者：T.Yagyu
'*                                      更新日：             更新者：
'*
'*  概要       ：SFR、SFRTシリーズ
'*
'************************************************************************************
Module KHPriceQ6

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strSw As String = ""

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            Dim bolC5Flag As Boolean

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
            objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
            objKtbnStrc.strcSelection.strOpSymbol(2)
            decOpAmount(UBound(decOpAmount)) = 1

            '2011/10/24 ADD RM1110032(11月VerUP:二次電池) START--->
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case ""
                    '2011/10/24 ADD RM1110032(11月VerUP:二次電池) <---END
                    'オプション加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        If (objKtbnStrc.strcSelection.strOpSymbol(5) = "R") Or (objKtbnStrc.strcSelection.strOpSymbol(5) = "L") Then
                            strSw = "S"
                        Else
                            strSw = "D"
                        End If
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(3) & objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                        strSw & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "2"
                    Dim intSu As Integer
                    ReDim strPriceDiv(0)
                    '基本価格キー用
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    'C5チェック
                    bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If

                    'オプション加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        If (objKtbnStrc.strcSelection.strOpSymbol(4) = "R") Or (objKtbnStrc.strcSelection.strOpSymbol(4) = "L") Then
                            strSw = "S"
                            intSu = 1
                        Else
                            strSw = "D"
                            intSu = 2
                        End If
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & CdCst.Sign.Hypen & _
                                                            objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                            strSw & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(1)
                        decOpAmount(UBound(decOpAmount)) = 1

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & CdCst.Sign.Hypen & _
                                                                "SW" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5)
                        decOpAmount(UBound(decOpAmount)) = intSu
                    End If

                    '二次電池加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) & CdCst.Sign.Hypen & _
                                                            objKtbnStrc.strcSelection.strOpSymbol(5)
                    decOpAmount(UBound(decOpAmount)) = 1
                    'strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5

            End Select
            '2011/10/24 ADD RM1110032(11月VerUP:二次電池) <---END

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module

