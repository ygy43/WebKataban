'************************************************************************************
'*  ProgramID  ：KHPriceO1
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/05   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：薬液マニホールド　ＧＡＭＤ
'*
'************************************************************************************
Module KHPriceO1

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim objKataban As New KHKataban
        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStationQty As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '連数設定
            intStationQty = CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)

            'ミックス判定
            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "X" Then
                For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                    If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
                        strOpArray = Split(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), CdCst.Sign.Comma)

                        If intLoopCnt <> 6 Then
                            '単体ブロック価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = KHKataban.fncHypenCut("AMD0*2A-" & _
                                                                                              strOpArray(4) & _
                                                                                              strOpArray(5) & _
                                                                                              strOpArray(6))
                            decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.intQuantity(intLoopCnt))
                        Else
                            'ベース価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = KHKataban.fncHypenCut("GAMD012A-BB-" & _
                                                                                              strOpArray(2) & _
                                                                                              strOpArray(3) & _
                                                                                              strOpArray(4) & _
                                                                                              strOpArray(5))
                            decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.intQuantity(intLoopCnt))
                        End If
                    End If
                Next
            Else
                'ベース価格キー
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "GAMD012A-BB-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                           intStationQty.ToString & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                '単体ブロック価格キー
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "AMD0*2A-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                decOpAmount(UBound(decOpAmount)) = CDec(intStationQty)
            End If

        Catch ex As Exception

            Throw ex

        Finally

            objKataban = Nothing

        End Try

    End Sub

End Module
