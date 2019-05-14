'************************************************************************************
'*  ProgramID  ：KHPrice01
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/12   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ブロックマニホールド　ＭＮ４ＫＢ１／ＭＮ４ＫＢ２
'*
'************************************************************************************
Module KHPrice01

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing, _
                                   Optional ByRef strCountryCd As String = Nothing, _
                                   Optional ByRef strOfficeCd As String = Nothing)

        Dim intLoopCnt As Integer
        Dim intStationQty1 As Integer = 0
        Dim intStationQty2 As Integer = 0

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            '仕様入力キー
            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                   objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                        Case CdCst.Manifold.InspReportJp.SelectValue, CdCst.Manifold.InspReportEn.SelectValue
                            '加算なし
                        Case Else
                            If Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "M" & objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                '数量のカウント(電磁弁付バルブブロック)
                                If intLoopCnt >= 9 And intLoopCnt <= 14 Then
                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = "1" Then
                                        intStationQty1 = intStationQty1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        intStationQty1 = intStationQty1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                    End If
                                Else
                                    intStationQty1 = intStationQty1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & _
                                                                           objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End If
                            If Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2) = "GW" Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Joint
                            End If
                    End Select
                End If
            Next
            'DINレール長さ加算価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & "DIN"
            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.decDinRailLength
            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail

            '数量セット
            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                Case "", "0"
                    intStationQty2 = 1
                Case Else
                    intStationQty2 = CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
            End Select

            'シリーズオプション加算
            For intLoopCnt = 3 To objKtbnStrc.strcSelection.strOpSymbol.Length - 1
                '2010/08/27 ADD RM0808112(異電圧対応) START--->
                'If intLoopCnt <> 7 Then
                '    If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim <> "" Then
                Dim isAdd As Boolean = False
                If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim <> "" Then
                    Select Case intLoopCnt
                        Case 7
                            'Do nothing
                        Case 8
                            '電圧

                            If KHKataban.fncVoltageIsStandard(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, _
                                                            strCountryCd, strOfficeCd) Then
                                isAdd = True
                            End If
                        Case Else
                            isAdd = True
                    End Select
                    If isAdd Then
                        '2010/08/27 ADD RM0808112(異電圧対応) <--- END

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                        Select Case True
                            Case Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2) = "80"
                                decOpAmount(UBound(decOpAmount)) = intStationQty1
                            Case Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2) = "10"
                                decOpAmount(UBound(decOpAmount)) = intStationQty2
                            Case Else
                                decOpAmount(UBound(decOpAmount)) = intStationQty2 * 2
                        End Select
                    End If
                End If
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
