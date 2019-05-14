'************************************************************************************
'*  ProgramID  ：KHPrice44
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/12   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：省配線ブロックマニホールド　ＭＮ４ＴＢ１／ＭＮ４ＴＢ２／ＭＮ４ＴＢＸ１２
'*
'************************************************************************************
Module KHPrice44

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim strStdVoltageFlag As String = CdCst.VoltageDiv.Standard
        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim strOpt As String
        Dim intIndex As Integer
        Dim intStationQty As Integer = 0
        Dim intQuantity As Integer = 0

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "MN4TB1"
                    strOpt = "1"
                Case Else
                    strOpt = "2"
            End Select
            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6, 1) = "X" Then
                intIndex = 7
            Else
                intIndex = 9
            End If
            intStationQty = CInt(objKtbnStrc.strcSelection.strOpSymbol(intIndex).Trim)
            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                   objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                        Case CdCst.Manifold.InspReportJp.Japanese, CdCst.Manifold.InspReportJp.English, _
                             CdCst.Manifold.InspReportEn.Japanese, CdCst.Manifold.InspReportEn.English
                            '加算なし
                        Case Else
                            Select Case intLoopCnt
                                Case 23 To 24
                                    'Case 24 To 25
                                    'ケーブル
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case Else
                                    If intLoopCnt > 3 And intLoopCnt < 10 Then
                                        If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = "-" Then
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & " ", "K ") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & " ", InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & " ", "K ") - 1)
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        Else
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-H") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-H") - 1) & "-L"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    Else
                                        If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6, 1) = "X" And _
                                           (intLoopCnt = 2 Or intLoopCnt = 3) Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "1" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                        Else
                                            If intLoopCnt > 17 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & CdCst.Sign.Hypen & _
                                                                                           objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & strOpt & CdCst.Sign.Hypen & _
                                                                                           objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                            End If
                                        End If
                                        If intLoopCnt = 1 And objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim = "T10" Then
                                            If intStationQty > 4 Then
                                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-20P"
                                            Else
                                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-10P"
                                            End If
                                        End If
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                            End Select

                            Select Case Left(strOpRefKataban(UBound(strOpRefKataban)), 8)
                                Case "N4TB-GWP"
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Joint
                            End Select
                            If (Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) = "N4TB1" Or _
                                Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) = "N4TB2") And _
                                Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) >= "0" And _
                                Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) <= "9" Then
                                If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = 1 Then
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                End If
                            End If
                    End Select
                End If
            Next

            'DINレール加算価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = "N4TB-BAA"
            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.decDinRailLength
            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail

            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6, 1) = "X" Then
                intIndex = 2
            Else
                intIndex = 4
            End If
            'オプション加算(手動装置加算価格)
            If Left(objKtbnStrc.strcSelection.strOpSymbol(intIndex).Trim, 1) <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & strOpt & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(intIndex).Trim
                decOpAmount(UBound(decOpAmount)) = intQuantity
            End If

            intIndex = intIndex + 1
            '表示・保護回路加算価格
            If objKtbnStrc.strcSelection.strOpSymbol(intIndex).Trim <> "L" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & strOpt & "-MINUS-L"
                decOpAmount(UBound(decOpAmount)) = intQuantity
            End If

            intIndex = intIndex + 2
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intIndex), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "K", "A"
                        'Case "", "K", "A"
                        'Case Else
                        '価格キー設定
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & strOpt & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = intStationQty
                End Select
            Next

            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6, 1) = "X" Then
                intIndex = 6
            Else
                intIndex = 8
            End If
            '切削油オプション加算価格
            If objKtbnStrc.strcSelection.strOpSymbol(intIndex).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & strOpt & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(intIndex).Trim
                decOpAmount(UBound(decOpAmount)) = intStationQty
            End If

            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6, 1) = "X" Then
                intIndex = 8
            Else
                intIndex = 10
            End If
            '電圧オプション加算価格
            If objKtbnStrc.strcSelection.strOpSymbol(intIndex).Trim <> "" Then
                strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                               objKtbnStrc.strcSelection.strOpSymbol(intIndex).Trim)
                Select Case strStdVoltageFlag
                    Case CdCst.VoltageDiv.Standard
                    Case CdCst.VoltageDiv.Options
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & strOpt & CdCst.Sign.Hypen & "OPT"
                        decOpAmount(UBound(decOpAmount)) = intQuantity
                    Case CdCst.VoltageDiv.Other
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & strOpt & CdCst.Sign.Hypen & "OTH"
                        decOpAmount(UBound(decOpAmount)) = intQuantity
                End Select
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
