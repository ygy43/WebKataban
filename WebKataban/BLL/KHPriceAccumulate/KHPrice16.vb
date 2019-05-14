'************************************************************************************
'*  ProgramID  ：KHPrice16
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/07   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：セルシリンダ　４Ｆ／Ｍ４Ｆ
'*
'************************************************************************************
Module KHPrice16

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strCountryCd As String = Nothing, _
                                   Optional ByRef strOfficeCd As String = Nothing)


        Dim strStdVoltageFlag As String = CdCst.VoltageDiv.Standard
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim strKataban1 As String = ""
        Dim strKataban2 As String = ""
        Dim intStationQty1 As Integer = 0
        Dim intStationQty2 As Integer = 0
        Dim intStationQty3 As Integer = 0
        Dim Hantei As String = ""

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            If objKtbnStrc.strcSelection.strKeyKataban = "M" Then
                '基本価格キー
                If objKtbnStrc.strcSelection.strSpecNo.Trim <> "" Then
                    If objKtbnStrc.strcSelection.strSpecNo.Trim = "52" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "A4F0" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                        intStationQty2 = CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                    Else
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                            '仕様有り
                            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                                If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)

                                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Contains("MP") Then

                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)

                                    ElseIf objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Contains("A4F") Then

                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 5) & "1"

                                    ElseIf objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Contains("4F0") Or _
                                        objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Contains("4F1") Then

                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4) & "1"

                                    ElseIf objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Contains("4F2") Or _
                                        objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Contains("4F3") Then

                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8)
                                            Case "C", "I"
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4) & "8"
                                            Case Else
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4) & "1"
                                        End Select

                                    ElseIf objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Contains("4F4") Or _
                                        objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Contains("4F5") Or _
                                        objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Contains("4F6") Or _
                                        objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Contains("4F7") Then

                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4) & "8"

                                    Else

                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)

                                    End If
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    '数量設定
                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 1) <> "M" Then
                                        intStationQty1 = intStationQty1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    Select Case True
                                        Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 2) = "01" Or _
                                             Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 1) = "1"
                                            intStationQty2 = intStationQty2 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 2) <> "0M" And _
                                             Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 1) <> "M"
                                            intStationQty2 = intStationQty2 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                    End Select
                                End If
                            Next
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 3) & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                            If Mid(strOpRefKataban(UBound(strOpRefKataban)), 4, 1) <> "M" Then
                                intStationQty1 = intStationQty1 + decOpAmount(UBound(decOpAmount))
                            End If
                            Select Case True
                                Case Mid(strOpRefKataban(UBound(strOpRefKataban)), 4, 2) = "01" Or _
                                     Mid(strOpRefKataban(UBound(strOpRefKataban)), 4, 1) = "1"
                                    intStationQty2 = intStationQty2 + decOpAmount(UBound(decOpAmount))
                                Case Mid(strOpRefKataban(UBound(strOpRefKataban)), 4, 2) <> "0M" And _
                                     Mid(strOpRefKataban(UBound(strOpRefKataban)), 4, 1) <> "M"
                                    intStationQty2 = intStationQty2 + decOpAmount(UBound(decOpAmount)) * 2
                            End Select
                        End If
                    End If
                Else
                    '仕様書無し
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        strKataban1 = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & _
                                      objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                      objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    Else
                        strKataban1 = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                      objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                      objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    End If
                    If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "6" Or _
                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "7" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "8" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "9" Then
                            strKataban1 = strKataban1 & CdCst.Sign.Hypen & _
                                          objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        End If
                    End If

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strKataban1
                    decOpAmount(UBound(decOpAmount)) = 1

                    '数量設定
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Then
                        intStationQty2 = 1
                    Else
                        intStationQty2 = 2
                    End If
                    intStationQty1 = 1
                End If

                '数量3セット
                If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Or _
                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) = "AM" Then
                    intStationQty3 = CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                Else
                    intStationQty3 = 1
                End If

                'シリーズ形番セット
                If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "A" Then
                    strKataban1 = "4F0"
                    strKataban2 = "M4F0"
                Else
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        strKataban1 = Trim(Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 29))
                        strKataban2 = Trim(objKtbnStrc.strcSelection.strSeriesKataban.Trim)
                    Else
                        strKataban1 = Trim(objKtbnStrc.strcSelection.strSeriesKataban.Trim)
                    End If
                End If

                '接続口径加算
                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "08Y" Or _
                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "10Y" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & "-Y"
                    decOpAmount(UBound(decOpAmount)) = intStationQty3
                End If

                '手動装置加算
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & CdCst.Sign.Hypen & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = intStationQty2
                    End Select
                Next

                '電線接続加算
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    decOpAmount(UBound(decOpAmount)) = intStationQty2
                End If

                'オプション加算
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else
                            If strOpArray(intLoopCnt).Trim = "H" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                            Else
                                If Hantei = "W" Then
                                    If strOpArray(intLoopCnt).Trim = "N" Or strOpArray(intLoopCnt).Trim = "NC" Or strOpArray(intLoopCnt).Trim = "NO" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim & "-W"

                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                    End If
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                End If
                            End If

                            If strOpArray(intLoopCnt).Trim = "W" Then Hantei = "W"

                            If strOpArray(intLoopCnt).Trim = "K" Then
                                decOpAmount(UBound(decOpAmount)) = intStationQty1
                            Else
                                If Left(objKtbnStrc.strcSelection.strSeriesKataban, 1) = "M" Or _
                                   Left(objKtbnStrc.strcSelection.strSeriesKataban, 2) = "AM" Or _
                                   strOpArray(intLoopCnt).Trim = "S" Or _
                                   strOpArray(intLoopCnt).Trim = "W" Then

                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case "N", "NC", "NO"
                                            decOpAmount(UBound(decOpAmount)) = intStationQty3
                                        Case Else
                                            decOpAmount(UBound(decOpAmount)) = intStationQty2
                                    End Select
                                Else
                                    decOpAmount(UBound(decOpAmount)) = intStationQty1
                                End If
                            End If
                    End Select
                Next

                '排気／取付方式加算
                If Left(objKtbnStrc.strcSelection.strSeriesKataban, 1) = "M" Or _
                   Left(objKtbnStrc.strcSelection.strSeriesKataban, 2) = "AM" Then
                    Select Case True
                        Case objKtbnStrc.strcSelection.strSeriesKataban = "M4F0" Or _
                             objKtbnStrc.strcSelection.strSeriesKataban = "AM4F0"
                            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "CL" Or _
                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "CU" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban2 & "8" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban2 & "8" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case objKtbnStrc.strcSelection.strSeriesKataban = "M4F1" And _
                             objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "I"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strKataban2 & "8" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case objKtbnStrc.strcSelection.strSeriesKataban = "M4F1" And _
                             objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strKataban2 & "8" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case objKtbnStrc.strcSelection.strSeriesKataban = "M4F2"
                            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "I" Or _
                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "C" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban2 & "8" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban2 & "8" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strKataban2 & "8" & objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                End If

                '電圧加算
                '2010/08/31 ADD RM0808112(異電圧対応) START--->
                Dim isVolChk As Boolean = True
                Select Case objKtbnStrc.strcSelection.strSeriesKataban
                    Case "4F0", "A4F0", "AM4F0", "M4F0"
                        isVolChk = False
                End Select
                '2010/08/31 ADD RM0808112(異電圧対応) <---END
                If Left(objKtbnStrc.strcSelection.strSeriesKataban, 1) = "M" Or _
                   Left(objKtbnStrc.strcSelection.strSeriesKataban, 2) = "AM" Then
                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                        '2010/08/31 MOD RM0808112(異電圧対応) START--->
                        If isVolChk Then
                            strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                                                           strCountryCd, strOfficeCd)
                        Else
                            strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        End If
                        'strStdVoltageFlag = KHUnitPrice.fncVoltageInfoGet(objKtbnStrc, _
                        '                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        '2010/08/31 MOD RM0808112(異電圧対応) <---END
                        Select Case strStdVoltageFlag
                            Case CdCst.VoltageDiv.Standard
                            Case CdCst.VoltageDiv.Options
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & "-OPT"
                                decOpAmount(UBound(decOpAmount)) = intStationQty2
                            Case CdCst.VoltageDiv.Other
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & "-OTH"
                                decOpAmount(UBound(decOpAmount)) = intStationQty2
                        End Select
                    End If
                Else
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                        '2010/08/31 MOD RM0808112(異電圧対応) START--->
                        If isVolChk Then
                            strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                                                           strCountryCd, strOfficeCd)
                        Else
                            strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                        End If
                        'strStdVoltageFlag = KHUnitPrice.fncVoltageInfoGet(objKtbnStrc, _
                        '                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                        '2010/08/31 MOD RM0808112(異電圧対応) <---END
                        Select Case strStdVoltageFlag
                            Case CdCst.VoltageDiv.Standard
                            Case CdCst.VoltageDiv.Options
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & "-OPT"
                                decOpAmount(UBound(decOpAmount)) = intStationQty2
                            Case CdCst.VoltageDiv.Other
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & "-OTH"
                                decOpAmount(UBound(decOpAmount)) = intStationQty2
                        End Select
                    End If
                End If
            Else
                '基本価格キー
                If objKtbnStrc.strcSelection.strSpecNo.Trim <> "" Then
                    If objKtbnStrc.strcSelection.strSpecNo.Trim = "52" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "A4F0" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                        intStationQty2 = CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                    Else
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                            '仕様有り
                            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                                If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    '数量設定
                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 1) <> "M" Then
                                        intStationQty1 = intStationQty1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    Select Case True
                                        Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 2) = "01" Or _
                                             Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 1) = "1"
                                            intStationQty2 = intStationQty2 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 2) <> "0M" And _
                                             Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 1) <> "M"
                                            intStationQty2 = intStationQty2 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                    End Select
                                End If
                            Next
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 3) & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                            If Mid(strOpRefKataban(UBound(strOpRefKataban)), 4, 1) <> "M" Then
                                intStationQty1 = intStationQty1 + decOpAmount(UBound(decOpAmount))
                            End If
                            Select Case True
                                Case Mid(strOpRefKataban(UBound(strOpRefKataban)), 4, 2) = "01" Or _
                                     Mid(strOpRefKataban(UBound(strOpRefKataban)), 4, 1) = "1"
                                    intStationQty2 = intStationQty2 + decOpAmount(UBound(decOpAmount))
                                Case Mid(strOpRefKataban(UBound(strOpRefKataban)), 4, 2) <> "0M" And _
                                     Mid(strOpRefKataban(UBound(strOpRefKataban)), 4, 1) <> "M"
                                    intStationQty2 = intStationQty2 + decOpAmount(UBound(decOpAmount)) * 2
                            End Select
                        End If
                    End If
                Else
                    '仕様書無し
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        strKataban1 = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & _
                                      objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                      objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    Else
                        strKataban1 = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                      objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                      objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    End If
                    If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "6" Or _
                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "7" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "8" Or _
                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "9" Then
                            strKataban1 = strKataban1 & CdCst.Sign.Hypen & _
                                          objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        End If
                    End If

                    'NAMUR規格対応品追加対応 2016/11/21 追加 松原
                    'シリーズ型番が「4F1」「4F3」の場合のみ以下の処理を実施
                    If objKtbnStrc.strcSelection.strSeriesKataban = "4F1" Or objKtbnStrc.strcSelection.strSeriesKataban = "4F3" Then

                        'NAMUR規格対応品のキー型番の場合のみ処理を実施する
                        If objKtbnStrc.strcSelection.strKeyKataban = "A" Or objKtbnStrc.strcSelection.strKeyKataban = "N" Or objKtbnStrc.strcSelection.strKeyKataban = "Z" Then
                            strKataban1 = strKataban1 & "-NM"
                        End If

                    End If

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strKataban1
                    decOpAmount(UBound(decOpAmount)) = 1

                    '数量設定
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Then
                        intStationQty2 = 1
                    Else
                        intStationQty2 = 2
                    End If
                    intStationQty1 = 1
                End If

                '数量3セット
                If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Or _
                   Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) = "AM" Then
                    intStationQty3 = CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                Else
                    intStationQty3 = 1
                End If

                'シリーズ形番セット
                If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "A" Then
                    strKataban1 = "4F0"
                    strKataban2 = "M4F0"
                Else
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        strKataban1 = Trim(Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 29))
                        strKataban2 = Trim(objKtbnStrc.strcSelection.strSeriesKataban.Trim)
                    Else
                        strKataban1 = Trim(objKtbnStrc.strcSelection.strSeriesKataban.Trim)
                    End If
                End If

                'NAMUR規格対応品追加対応 2016/11/21 追加 松原
                'シリーズ型番が「4F1」「4F3」の場合のみ以下の処理を実施
                If objKtbnStrc.strcSelection.strSeriesKataban = "4F1" Or objKtbnStrc.strcSelection.strSeriesKataban = "4F3" Then

                    'NAMUR規格対応品のキー型番の場合のみ処理を実施する
                    If objKtbnStrc.strcSelection.strKeyKataban = "A" Or objKtbnStrc.strcSelection.strKeyKataban = "N" Or objKtbnStrc.strcSelection.strKeyKataban = "Z" Then
                        strKataban1 = strKataban1 & "-NM"
                    End If

                End If

                '接続口径加算
                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "08Y" Or _
                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "10Y" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & "-Y"
                    decOpAmount(UBound(decOpAmount)) = intStationQty3
                End If

                '手動装置加算
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & CdCst.Sign.Hypen & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = intStationQty2
                    End Select
                Next

                '電線接続加算
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    decOpAmount(UBound(decOpAmount)) = intStationQty2
                End If

                'オプション加算
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else
                            If strOpArray(intLoopCnt).Trim = "H" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                            Else
                                If Hantei = "W" Then
                                    If strOpArray(intLoopCnt).Trim = "N" Or strOpArray(intLoopCnt).Trim = "NC" Or strOpArray(intLoopCnt).Trim = "NO" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim & "-W"

                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                    End If
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                End If
                            End If

                            If strOpArray(intLoopCnt).Trim = "W" Then Hantei = "W"

                            If strOpArray(intLoopCnt).Trim = "K" Then
                                decOpAmount(UBound(decOpAmount)) = intStationQty1
                            Else
                                If Left(objKtbnStrc.strcSelection.strSeriesKataban, 1) = "M" Or _
                                   Left(objKtbnStrc.strcSelection.strSeriesKataban, 2) = "AM" Or _
                                   strOpArray(intLoopCnt).Trim = "S" Or _
                                   strOpArray(intLoopCnt).Trim = "W" Then

                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case "N", "NC", "NO"
                                            decOpAmount(UBound(decOpAmount)) = intStationQty3
                                        Case Else
                                            decOpAmount(UBound(decOpAmount)) = intStationQty2
                                    End Select
                                Else
                                    decOpAmount(UBound(decOpAmount)) = intStationQty1
                                End If
                            End If
                    End Select
                Next

                '排気／取付方式加算
                If Left(objKtbnStrc.strcSelection.strSeriesKataban, 1) = "M" Or _
                   Left(objKtbnStrc.strcSelection.strSeriesKataban, 2) = "AM" Then
                    Select Case True
                        Case objKtbnStrc.strcSelection.strSeriesKataban = "M4F0" Or _
                             objKtbnStrc.strcSelection.strSeriesKataban = "AM4F0"
                            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "CL" Or _
                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "CU" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban2 & "8" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban2 & "8" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case objKtbnStrc.strcSelection.strSeriesKataban = "M4F1" And _
                             objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "I"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strKataban2 & "8" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case objKtbnStrc.strcSelection.strSeriesKataban = "M4F1" And _
                             objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strKataban2 & "8" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case objKtbnStrc.strcSelection.strSeriesKataban = "M4F2"
                            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "I" Or _
                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "C" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban2 & "8" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban2 & "8" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strKataban2 & "8" & objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                End If


                'シリーズ型番を再セット 2016/11/26 追加 松原
                'シリーズ形番セット
                If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "A" Then
                    strKataban1 = "4F0"
                    strKataban2 = "M4F0"
                Else
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        strKataban1 = Trim(Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 29))
                        strKataban2 = Trim(objKtbnStrc.strcSelection.strSeriesKataban.Trim)
                    Else
                        strKataban1 = Trim(objKtbnStrc.strcSelection.strSeriesKataban.Trim)
                    End If
                End If

                'NAMUR規格対応品追加対応 2016/11/26 追加 松原
                'シリーズ型番が「4F～」の場合のみ以下の処理を実施
                If Mid(objKtbnStrc.strcSelection.strSeriesKataban, 1, 2) = "4F" Then

                    'NAMUR規格対応品のキー型番の場合のみ処理を実施する
                    If objKtbnStrc.strcSelection.strKeyKataban = "A" Then
                        strKataban1 = strKataban1 & "-NM"
                    End If

                End If

                '電圧加算
                '2010/08/31 ADD RM0808112(異電圧対応) START--->
                Dim isVolChk As Boolean = True
                Select Case objKtbnStrc.strcSelection.strSeriesKataban
                    Case "4F0", "A4F0", "AM4F0", "M4F0"
                        isVolChk = False
                End Select
                '2010/08/31 ADD RM0808112(異電圧対応) <---END
                If Left(objKtbnStrc.strcSelection.strSeriesKataban, 1) = "M" Or _
                   Left(objKtbnStrc.strcSelection.strSeriesKataban, 2) = "AM" Then
                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                        '2010/08/31 MOD RM0808112(異電圧対応) START--->
                        If isVolChk Then
                            strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                                                           strCountryCd, strOfficeCd)
                        Else
                            strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        End If
                        'strStdVoltageFlag = KHUnitPrice.fncVoltageInfoGet(objKtbnStrc, _
                        '                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        '2010/08/31 MOD RM0808112(異電圧対応) <---END
                        Select Case strStdVoltageFlag
                            Case CdCst.VoltageDiv.Standard
                            Case CdCst.VoltageDiv.Options
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & "-OPT"
                                decOpAmount(UBound(decOpAmount)) = intStationQty2
                            Case CdCst.VoltageDiv.Other
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & "-OTH"
                                decOpAmount(UBound(decOpAmount)) = intStationQty2
                        End Select
                    End If
                Else
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                        '2010/08/31 MOD RM0808112(異電圧対応) START--->
                        If isVolChk Then
                            strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                                                           strCountryCd, strOfficeCd)
                        Else
                            strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                        End If
                        'strStdVoltageFlag = KHUnitPrice.fncVoltageInfoGet(objKtbnStrc, _
                        '                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                        '2010/08/31 MOD RM0808112(異電圧対応) <---END
                        Select Case strStdVoltageFlag
                            Case CdCst.VoltageDiv.Standard
                            Case CdCst.VoltageDiv.Options
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & "-OPT"
                                decOpAmount(UBound(decOpAmount)) = intStationQty2
                            Case CdCst.VoltageDiv.Other
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strKataban1 & "-OTH"
                                decOpAmount(UBound(decOpAmount)) = intStationQty2
                        End Select
                    End If
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
