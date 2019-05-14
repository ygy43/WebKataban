'************************************************************************************
'*  ProgramID  ：KHPriceR6
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2012/04/09   作成者：Y.Tachi
'*
'*  概要       ：スプール位置検出機能付３ポート電磁弁(ＳＮＰ)
'*
'************************************************************************************
Module KHPriceR6

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOptionKataban As String = ""

        Try

            'RM1807020_機種追加
            Select Case objKtbnStrc.strcSelection.strSeriesKataban

                Case "SNP"
                    '配列定義
                    ReDim strOpRefKataban(0)
                    ReDim decOpAmount(0)


                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-02GL"
                    decOpAmount(UBound(decOpAmount)) = 1

                    'コイルオプション価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "2G" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        '↓RM1301005 2013/01/11 Y.Tachi
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "2" Then
                            '連数が２連の場合は２倍加算
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If

                    'リミットスイッチ価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "L" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        '↓RM1301005 2013/01/11 Y.Tachi
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "2" Then
                            '連数が２連の場合は２倍加算
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If

                    'ブラケット価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'サイレンサ価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        '↓RM1301005 2013/01/11 Y.Tachi
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "2" Then
                            '連数が２連の場合は２倍加算
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If

                    '二次電池
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "4" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "SNS"

                    '配列定義
                    ReDim strOpRefKataban(0)
                    ReDim decOpAmount(0)


                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '電線接続加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(1).Trim

                    'セーフティリミットスイッチ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SW"
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SW-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    End If
                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(1).Trim

                    'パイロット方式加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    End If

                    '流れ方向加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    End If

                    'サイレンサ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    End If

            End Select

        Catch ex As Exception

            Throw ex

        Finally
        End Try

    End Sub

End Module