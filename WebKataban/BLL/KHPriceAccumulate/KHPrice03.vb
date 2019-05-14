'************************************************************************************
'*  ProgramID  ：KHPrice03
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/12   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：直動式２ポート弁　ＡＢ／ＡＧ
'*
'*  ・受付No：RM0907070  二次電池対応機器　
'*                                      更新日：2009/09/08   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　KHOptionCtl.vb
'*                                      更新日：2010/02/22   更新者：Y.Miura
'*  ・受付No：RM0808112  異電圧対応
'*                                      更新日：2010/08/11   更新者：Y.Miura
'************************************************************************************
Module KHPrice03

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing, _
                                   Optional ByRef strCountryCd As String = Nothing, _
                                   Optional ByRef strOfficeCd As String = Nothing)


        Dim strStdVoltageFlag As String = CdCst.VoltageDiv.Standard
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim strPort As String
        Dim bolScrew As Boolean
        Dim intStationQty As Integer = 0
        Dim bolOptionZ As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            '基本価格キー
            If InStr(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, "G") <> 0 Or _
               InStr(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, "N") <> 0 Then
                strPort = "0" & Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1)
                bolScrew = True
            Else
                strPort = Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2)
                bolScrew = False
            End If
            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "AB71" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           strPort & CdCst.Sign.Hypen & _
                                                           Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 2)
                decOpAmount(UBound(decOpAmount)) = 1
            Else

                If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) <> "" And _
                   Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) <> "0" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               strPort & CdCst.Sign.Hypen & _
                                                               Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) & CdCst.Sign.Hypen & _
                                                               Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1)
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               strPort & CdCst.Sign.Hypen & _
                                                               Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1)
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            'コイルハウジング加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "AB71" Then
                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) <> "B" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_03" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            Else
                If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "2" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        'RM0907070 2009/09/08 Y.Miura 二次電池対応
                        'strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_04" & _
                        '                                           Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2) & _
                        '                                           Left(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, 2)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_04" & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2) & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_04" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            End If

            '手動操作加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "AB71" Then
                If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_04" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            Else
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_05" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            '取付板加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "AB71" Then
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_05" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            Else
                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_06" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            'ケーブルグランド・コンジット加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "AB71" Then
                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_06" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            Else
                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_07" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            'オプション加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.Trim <> "AB71" Then
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case "S"
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2) = "" Or _
                               Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2) = "00" Or _
                               Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2) = "3A" Or _
                               Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2) = "4A" Or _
                               Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2) = "6C" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_08S0"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_08" & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_08" & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Next

                'オプション２　Ｐ４加算
                'RM0907070 2009/09/08 Y.Miura　二次電池対応
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(9), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else       '"P4", "P40"を含む
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                Next

            End If

            '電圧加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.Trim <> "AB71" Then
                'RM0907070 2009/09/08 Y.Miura 二次電池対応
                'strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                '                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                'RM0808112　異電圧対応
                'strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                '                                               objKtbnStrc.strcSelection.strOpSymbol(10).Trim)
                strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                               objKtbnStrc.strcSelection.strOpSymbol(10).Trim, strCountryCd, strOfficeCd)
                Select Case strStdVoltageFlag
                    Case CdCst.VoltageDiv.Standard
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        'RM0907070 2009/09/08 Y.Miura 二次電池対応
                        'strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_09" & _
                        '                                           Left(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, 2)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "_09" & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            End If

            'ねじ加算価格キー
            'If bolScrew Then
            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            '    strOpRefKataban(UBound(strOpRefKataban)) = "MULTI-SCREW-" & Right(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1)
            '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Screw
            '    Select Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2)
            '        Case "AB"
            '            decOpAmount(UBound(decOpAmount)) = 2
            '        Case "AG"
            '            decOpAmount(UBound(decOpAmount)) = 3
            '    End Select
            'End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
