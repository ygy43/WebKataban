'************************************************************************************
'*  ProgramID  ：KHPrice52
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/08   作成者：NII K.Sudoh
'*
'*  概要       ：マイクロゾール　Ｐ５１／Ｍ５１／Ｂ５１／Ｂ＊Ｐ５１／Ｗ＊Ｐ５１／Ｎ＊Ｐ５１
'*
'*  更新履歴   ：                       更新日：2007/05/24   更新者：NII A.Takahashi
'*               ・Ｂ＊Ｐ５１において、電磁弁単価の形番がただしく生成されていないのを修正
'************************************************************************************
Module KHPrice52

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim strStdVoltageFlag As String = CdCst.VoltageDiv.Standard
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim intQuantity As Integer = 0
        Dim intIndex As Integer = 0
        Dim strSiyItemCd() As String
        Dim decSiyItemNb() As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            ReDim strSiyItemCd(0)
            ReDim decSiyItemNb(0)

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "B"
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "8" Then
                        For intLoopCnt = objKtbnStrc.strcSelection.intQuantity.Length - 1 To 1 Step -1
                            If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                                If InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), "MP") = 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                               Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4, 1) & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim & "-SP" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    Exit For
                                End If
                            End If
                        Next

                        For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                            ReDim Preserve strSiyItemCd(UBound(strSiyItemCd) + 1)
                            ReDim Preserve decSiyItemNb(UBound(decSiyItemNb) + 1)
                            strSiyItemCd(UBound(strSiyItemCd)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                            decSiyItemNb(UBound(decSiyItemNb)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                        Next

                        For intLoopCnt = 1 To strSiyItemCd.Length - 1
                            If strSiyItemCd(intLoopCnt).Trim <> "" And _
                               decSiyItemNb(intLoopCnt) <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strSiyItemCd(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = decSiyItemNb(intLoopCnt)
                                If Mid(strSiyItemCd(intLoopCnt).Trim, 6, 2) <> "MP" Then
                                    intQuantity = intQuantity + decSiyItemNb(intLoopCnt)
                                End If
                            End If
                        Next
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(9).Trim & "-SP" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(1).Trim)
                        intQuantity = decOpAmount(UBound(decOpAmount))
                    End If

                    '手動装置加算～金具加算
                    For intLoopCnt = 5 To 8
                        If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "P51"
                            If intIndex = 7 And objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "P51" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            End If
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        End If
                    Next

                    'その他のオプション加算
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = ""
                                If strOpArray(intLoopCnt).Trim = "L" Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                End If
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "P51" & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                If strOpArray(intLoopCnt).Trim = "L" Then
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    decOpAmount(UBound(decOpAmount)) = intQuantity
                                End If
                        End Select
                    Next

                    '電圧加算
                    If objKtbnStrc.strcSelection.strOpElementDiv(11).Trim = CdCst.ElementDiv.Voltage Then
                        strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(11).Trim)
                        Select Case strStdVoltageFlag
                            Case CdCst.VoltageDiv.Standard
                            Case CdCst.VoltageDiv.Options
                                If InStr(objKtbnStrc.strcSelection.strOpSymbol(10), "L") > 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & "P51-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim & "P51-OPT"
                                    decOpAmount(UBound(decOpAmount)) = intQuantity
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "P51-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim & "P51-OPT"
                                    decOpAmount(UBound(decOpAmount)) = intQuantity
                                End If
                            Case Else
                                If InStr(objKtbnStrc.strcSelection.strOpSymbol(10), "L") > 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & "P51-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim & "P51-OTH"
                                    decOpAmount(UBound(decOpAmount)) = intQuantity
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "P51-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim & "P51-OTH"
                                    decOpAmount(UBound(decOpAmount)) = intQuantity
                                End If
                        End Select
                    End If
                Case "N"
                    '数量設定(連数)
                    intQuantity = CInt(objKtbnStrc.strcSelection.strOpSymbol(1).Trim)

                    For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                        If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                           objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
                            Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                Case CdCst.Manifold.InspReportJp.SelectValue, CdCst.Manifold.InspReportEn.SelectValue
                                    '加算なし
                                Case Else
                                    ReDim Preserve strSiyItemCd(UBound(strSiyItemCd) + 1)
                                    ReDim Preserve decSiyItemNb(UBound(decSiyItemNb) + 1)
                                    decSiyItemNb(UBound(decSiyItemNb)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    If intLoopCnt < 3 Or intLoopCnt > 12 Then
                                        strSiyItemCd(UBound(strSiyItemCd)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "51-" & _
                                                                             objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    Else
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") <> 0 Then
                                            strSiyItemCd(UBound(strSiyItemCd)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1)
                                            If strSiyItemCd(UBound(strSiyItemCd) - 1) = strSiyItemCd(UBound(strSiyItemCd)) Then
                                                decSiyItemNb(UBound(decSiyItemNb) - 1) = decSiyItemNb(UBound(decSiyItemNb) - 1) + decSiyItemNb(UBound(decSiyItemNb))
                                                strSiyItemCd(UBound(strSiyItemCd)) = ""
                                                decSiyItemNb(UBound(decSiyItemNb)) = 0
                                            End If
                                        End If
                                    End If
                            End Select
                        End If
                    Next

                    For intLoopCnt = 1 To strSiyItemCd.Length - 1
                        If strSiyItemCd(intLoopCnt).Trim <> "" And decSiyItemNb(intLoopCnt) > 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strSiyItemCd(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = decSiyItemNb(intLoopCnt)
                        End If
                    Next

                    'DINレール加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "N51-BAA"
                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.decDinRailLength
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail

                    '手動装置加算～金具加算
                    For intLoopCnt = 5 To 8
                        If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "51" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        End If
                    Next

                    'その他のオプション加算
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "51" & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intQuantity
                        End Select
                    Next

                    '電圧加算
                    If objKtbnStrc.strcSelection.strOpElementDiv(11).Trim = CdCst.ElementDiv.Voltage Then
                        strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(11).Trim)
                        Select Case strStdVoltageFlag
                            Case CdCst.VoltageDiv.Standard
                            Case CdCst.VoltageDiv.Options
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "51" & "-OPT"
                                decOpAmount(UBound(decOpAmount)) = intQuantity
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "51" & "-OTH"
                                decOpAmount(UBound(decOpAmount)) = intQuantity
                        End Select
                    End If
                Case "W"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim & "-SP" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(1).Trim)
                    intQuantity = decOpAmount(UBound(decOpAmount))

                    '手動装置加算～金具加算
                    For intLoopCnt = 5 To 8
                        If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "P51" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        End If
                    Next

                    'その他のオプション加算
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "P51" & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intQuantity
                        End Select
                    Next

                    '電圧加算
                    If objKtbnStrc.strcSelection.strOpElementDiv(11).Trim = CdCst.ElementDiv.Voltage Then
                        strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(11).Trim)
                        Select Case strStdVoltageFlag
                            Case CdCst.VoltageDiv.Standard
                            Case CdCst.VoltageDiv.Options
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "P51-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim & "P51-OPT"
                                decOpAmount(UBound(decOpAmount)) = intQuantity
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "P51-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim & "P51-OTH"
                                decOpAmount(UBound(decOpAmount)) = intQuantity
                        End Select
                    End If
                Case Else
                    '基本価格
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '手動装置加算～金具加算
                    For intLoopCnt = 3 To 7
                        If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim
                            If intLoopCnt = 7 Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            End If
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Next

                    'その他のオプション加算
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                If strOpArray(intLoopCnt).Trim = "L" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & "KANAGU"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                    Next

                    '電圧加算
                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                        strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        Select Case strStdVoltageFlag
                            Case CdCst.VoltageDiv.Standard
                            Case CdCst.VoltageDiv.Options
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OPT"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OTH"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
