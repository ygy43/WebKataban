'************************************************************************************
'*  ProgramID  ：KHPriceD2
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/21   作成者：NII K.Sudoh
'*
'*  概要       ：スーパーロッドレスシリンダ　ＳＲＬ２－Ｊ／ＧＱ
'*               スーパーロッドレスシリンダ　ＳＲＬ３
'*
'*  更新履歴   ：                       更新日：2008/01/08   更新者：NII A.Takahashi
'*               ・SRL3の単価ロジック追加
'*  ・受付No：RM0907070  二次電池対応機器　SRL3
'*                                      更新日：2009/08/21   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'************************************************************************************
Module KHPriceD2

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer
        Dim bolC5Flag As Boolean

        Dim bolOptionI As Boolean
        Dim bolOptionY As Boolean
        Dim bolOptionP4 As Boolean = False          'RM0907070 2009/08/21 Y.Miura　二次電池対応
        Dim strOptionP4 As String = String.Empty

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)                        'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応

            '初期設定
            bolC5Flag = False
            bolOptionI = False
            bolOptionY = False

            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "SRL3"
                    'C5チェック
                    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                    'bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc)
                    bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

                    'RM0907070 2009/08/21 Y.Miura　二次電池対応
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""

                            Case "P4", "P40"
                                bolOptionP4 = True
                                strOptionP4 = strOpArray(intLoopCnt).Trim
                        End Select
                    Next

                    'ストローク取得
                    intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim))

                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        'RM0907070 2009/08/21 Y.Miura　二次電池対応
                        'Case "", "Q"
                        Case "", "Q", "4", "R", "F"
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                       intStroke.ToString
                        Case Else
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                                       Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                       intStroke.ToString
                    End Select
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If

                    'バリエーション「Q」(落下防止)加算価格キー
                    If InStr(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, "Q") > 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                                   "Q" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    End If

                    '支持形式加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "00" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'スイッチ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                                   "SW" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)

                        'リード線長さ加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Case "M0H", "M0V", "M5H", "M5V", "M2H", "M2V", "M2WV", "M3H", "M3V", "M3WV", "M3PH", "M3PV", _
                                     "T2WH", "T2WV", "T2YH", "T2YV", "T3WH", "T3WV", "T3YH", "T3YV", "T2YLH", "T2YLV", "T3YLH", "T3YLV", _
                                     "T1H", "T1V", "T2H", "T2V", "T3H", "T3V", "T3PH", "T3PV", "T0H", "T0V", "T5H", "T5V", "T8H", "T8V", _
                                     "T0HF", "T0VF", "T0HM", "T0VM", "T0HU", "T0VU", "T2HF", "T2VF", "T2HM", "T2VM", "T2HU", "T2VU", _
                                     "T2WHF", "T2WVF", "T2WHM", "T2WVM", "T2WHU", "T2WVU", "T3HF", "T3VF", "T3PHF", "T3PVF", "T3WHF", "T3WVF"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                                               "SWLW(1)" & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                Case "T2YD"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                                               "SWLW(2)" & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                Case "T2YDT"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                                               "SWLW(3)" & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                            End Select
                        End If

                        'RM0907070 2009/08/21 Y.Miura　二次電池対応
                        'P4スイッチ加算
                        If bolOptionP4 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & "-SW-P4"
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        End If

                    End If

                    'オプション加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                Select Case Left(strOpArray(intLoopCnt).Trim, 1)
                                    Case "L", "N"
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                                                   Left(strOpArray(intLoopCnt).Trim, 1) & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        decOpAmount(UBound(decOpAmount)) = Right(strOpArray(intLoopCnt).Trim, 1)
                                        If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case "A"
                                        If InStr(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, "Q") > 0 Then
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                                                       "Q" & CdCst.Sign.Hypen & _
                                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Else
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        End If
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If

                                        'RM0907070 2009/08/21 Y.Miura　二次電池対応
                                        Select Case strOpArray(intLoopCnt).Trim
                                            Case "A", "A1", "A2"    'ショックキラー付
                                                If strOptionP4 <> "" Then 'ショックキラーのＰ４加算
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                                                               Left(strOpArray(intLoopCnt).Trim, 1) & CdCst.Sign.Hypen & _
                                                                                               strOptionP4 & CdCst.Sign.Hypen & _
                                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim

                                                    Select Case strOpArray(intLoopCnt).Trim
                                                        Case "A"    '２ヶ付
                                                            decOpAmount(UBound(decOpAmount)) = 2
                                                        Case Else   '１ヶ付
                                                            decOpAmount(UBound(decOpAmount)) = 1
                                                    End Select
                                                End If
                                            Case "A3"               'ショックキラー無し
                                            Case Else
                                        End Select

                                    Case Else
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                        End Select
                    Next

                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "F", "H"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & "-OP-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(11).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                    End Select

                Case Else
                    'ストローク取得
                    intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                          CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim), _
                                                          CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim))

                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 6) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If

                    'バリエーション「Q」(落下防止)加算価格キー
                    If Mid(objKtbnStrc.strcSelection.strSeriesKataban, 7, 1) = "Q" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & "Q" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    End If

                    '支持形式加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "00" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'スイッチ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)

                        'リード線長さ加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                            Select Case Mid(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 4, 1)
                                Case "F", "M"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & "FM"
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                                Case "D"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                                Case "L"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & "L"
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                            End Select
                        End If
                    End If

                    'オプション加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                If Left(strOpArray(intLoopCnt).Trim, 1) = "L" Or Left(strOpArray(intLoopCnt).Trim, 1) = "N" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                               Left(strOpArray(intLoopCnt).Trim, 1) & "1" & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    decOpAmount(UBound(decOpAmount)) = Val(Mid(strOpArray(intLoopCnt).Trim, 2, 1))

                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1

                                    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case "P4", "P40"
                                        Case Else
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                                End If
                        End Select
                    Next
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
