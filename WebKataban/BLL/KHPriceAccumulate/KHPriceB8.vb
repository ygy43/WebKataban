'************************************************************************************
'*  ProgramID  ：KHPriceB8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/06   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＥＶＴ／ＭＥＶＴ
'*
'************************************************************************************
Module KHPriceB8

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "MEVT"
                    For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                        'オプション形番が存在し且つ使用数が0より大きい場合
                        If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                           objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
                            Select Case intLoopCnt
                                Case 1 To 1
                                    'EVT
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'クリーン仕様加算価格キー
                                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "P70" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6) & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                Case 2 To 4
                                    '電装・給排気ブロック
                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 8, 1) = "R" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 8)
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 9)
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                Case 5 To 6
                                    'エンドブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case 7 To 9
                                    'ブランクプラグ・サイレンサ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "EVT" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End Select
                        End If
                    Next

                    'DINレール
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "EVT-BAA"
                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.decDinRailLength
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail

                    'リティナ(左)加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "EVT-HL"
                    decOpAmount(UBound(decOpAmount)) = 1

                    'リティナ(右)加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "EVT-HR"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "EVT"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'クリーン仕様加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "P70" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
