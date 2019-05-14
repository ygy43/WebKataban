'************************************************************************************
'*  ProgramID  ：KHPriceE5
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/26   作成者：NII K.Sudoh
'*
'*  概要       ：落下防止付クランプシリンダ　ＵＣＡＣ／ＵＣＡＣ－Ｌ２／ＵＣＡＣ２／ＵＣＡＣ２－Ｌ２
'*
'*  更新履歴   ：                       更新日：2007/05/16   更新者：NII A.Takahashi
'*               ・T2W/T3Wスイッチ追加に伴い、リード線加算ロジック部を修正
'*  ・受付No：RM0811133  UCAC2新発売
'*                                      更新日：2009/07/28   更新者：Y.Miura
'*  ・受付No：RM1001018　スイッチT2YDUをC5扱いとする
'*                                      更新日：2010/01/18   更新者：Y.Miura
'*  ・受付No：RM1003086　タイロッド取付位置追加   2010/03/26 Y.Miura        
'************************************************************************************
Module KHPriceE5

    'Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
    '                               ByRef strOpRefKataban() As String, _
    '                               ByRef decOpAmount() As Decimal)
    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer
        Dim bolC5Flag As Boolean    'RM0811133 2009/07/28 Y.Miura 追加

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)    'RM0811133 2009/07/28 Y.Miura 追加

            'RM0811133 2009/07/28 Y.Miura　↓↓
            'シリーズ形番の第1ハイフン前を取得
            Dim strHySeriesKataban As String = objKtbnStrc.strcSelection.strSeriesKataban.Trim
            If InStr(strHySeriesKataban, "-") > 0 Then
                strHySeriesKataban = strHySeriesKataban.Substring(0, InStr(strHySeriesKataban, "-") - 1)
            End If
            '要素位置の設定
            'UCAC2は3番目の要素『配管ねじ種類』が存在するのでstrOpSymbol(3)以降はプラス1する
            'RM1003086 2010/03/26 Y.Miura 
            'タイロッド取付位置追加に伴い、UCAC2の『付属品』以降はさらにプラス1する
            Dim intOpt As Integer
            Dim intOpt2 As Integer
            Select Case strHySeriesKataban
                Case "UCAC"
                    intOpt = 0
                    intOpt2 = 0
                Case "UCAC2"
                    intOpt = 1
                    intOpt2 = 1
            End Select

            'C5チェック
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

            'ストローク取得
            'intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc.strcSelection.strSeriesKataban, _
            '                                      objKtbnStrc.strcSelection.strKeyKataban, _
            '                                      CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim), _
            '                                      CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim))
            intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                     CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim), _
                                                     CInt(objKtbnStrc.strcSelection.strOpSymbol(4 + intOpt).Trim))
            'RM0811133 2009/07/28 Y.Miura　↑↑

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                       intStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1
            'RM0811133 2009/07/28 Y.Miura
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
            End If

            'クレビス幅加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "AL", "BL", "C", "CL"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    'RM0811133 2009/07/28 Y.Miura
                    'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                    '                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    'RM0811133 2009/07/28 Y.Miura
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
            End Select

            'スイッチ加算価格キー
            'RM0811133 2009/07/28 Y.Miura
            'If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
            If objKtbnStrc.strcSelection.strOpSymbol(7 + intOpt).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                'RM0811133 2009/07/28 Y.Miura ↓↓
                'strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                '                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                Select Case strHySeriesKataban
                    Case "UCAC"
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7 + intOpt).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                    Case "UCAC2"
                        '↓RM1309001 2013/09/02 追加
                        If objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "Z" Or objKtbnStrc.strcSelection.strOpSymbol(12).Trim = "Z" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-Z-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7 + intOpt).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                        Else

                            strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7 + intOpt).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
                'RM0811133 2009/07/28 Y.Miura ↑↑

                'リード線の長さ加算価格キー
                'RM0811133 2009/07/28 Y.Miura
                'If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                If objKtbnStrc.strcSelection.strOpSymbol(8 + intOpt).Trim <> "" Then

                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                        'RM0811133 2009/07/28 Y.Miura ↓↓
                        'Case "UCAC"
                        '    Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        '        Case "T0H", "T2H", "T3H", "T5H", "T1H", "T8H"
                        '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-T*H-" & _
                        '                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                        '            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        '        Case "T0V", "T2V", "T3V", "T5V", "T1V", "T8V"
                        '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-T*V-" & _
                        '                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                        '            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        '        Case "T2YH", "T3YH", "T2WH", "T3WH"
                        '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-T*YH-" & _
                        '                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                        '            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        '        Case "T2YV", "T3YV", "T2WV", "T3WV"
                        '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-T*YV-" & _
                        '                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                        '            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        '        Case Else
                        '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                        '                                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                        '                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                        '            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        '    End Select
                        'Case "UCAC-L2"
                        '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                        '                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                        '                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                        '    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        Case "UCAC-L2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7 + intOpt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8 + intOpt).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                            '↓RM1309001 2013/09/02 追加
                        Case "UCAC2"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7 + intOpt).Trim
                                Case "T0H", "T2H", "T3H", "T5H", "T1H", "T8H", "T3PH", "T2JH"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*H-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                                Case "T0V", "T2V", "T3V", "T5V", "T1V", "T8V", "T3PV", "T2JV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*V-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                                Case "T2YH", "T3YH", "T2WH", "T3WH"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*YH-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                                Case "T2YV", "T3YV", "T2WV", "T3WV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*YV-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(7 + intOpt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                    End If
                            End Select
                            '↑RM1309001 2013/09/02 追加
                        Case Else
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7 + intOpt).Trim
                                Case "T0H", "T2H", "T3H", "T5H", "T1H", "T8H"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*H-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                                Case "T0V", "T2V", "T3V", "T5V", "T1V", "T8V"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*V-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                                Case "T2YH", "T3YH", "T2WH", "T3WH"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*YH-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                                Case "T2YV", "T3YV", "T2WV", "T3WV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-T*YV-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(7 + intOpt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8 + intOpt).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                            End Select
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                            'RM0811133 2009/07/28 Y.Miura ↑↑
                    End Select
                End If

                '取付用タイロッド加算価格キー
                'RM0811133 2009/07/28 Y.Miura ↓↓
                'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-TIEROD"
                Select Case strHySeriesKataban
                    Case "UCAC"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-TIEROD"
                    Case "UCAC2"
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "R", "S"
                                'タイロット加算なし
                            Case Else
                                '↓RM1309001 2013/09/02 追加
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                If objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "Z" Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-Z-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                                Else
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & CdCst.Sign.Hypen & "TIEROD" & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               intStroke.ToString
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                End Select
                '↓RM1309001 2013/09/02 追加
                ''RM0811133 2009/07/28 Y.Miura ↑↑
                'decOpAmount(UBound(decOpAmount)) = 1
                ''RM0811133 2009/07/28 Y.Miura
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If


            '取付用タイロッド加算価格キー(スイッチなし時) RM1003086 2010/03/26 Y.Miura 追加
            Select Case strHySeriesKataban
                Case "UCAC2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "S"
                            'タイロット加算なし
                        Case Else
                            If objKtbnStrc.strcSelection.strOpSymbol(12).Trim <> "" Then
                                '2013/09/02 追加
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim = "Z" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & "-Z-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9 + intOpt).Trim)
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & CdCst.Sign.Hypen & "TIEROD" & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               intStroke.ToString
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            End If
                    End Select
            End Select

            '付属品加算価格キー
            'RM1003086 2010/03/26 Y.Miura 変更
            'strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(11 + intOpt), CdCst.Sign.Delimiter.Comma)
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(11 + intOpt + intOpt2), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        'strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                        '                                           strOpArray(intLoopCnt).Trim
                        Select Case strHySeriesKataban
                            Case "UCAC"
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                            Case "UCAC2"
                                strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                        End Select
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0811133 2009/07/28 Y.Miura
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                End Select
            Next

            'スズキ向け特注
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "UCAC2", "UCAC2-L2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "S"
                            'If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            '    strOpRefKataban(UBound(strOpRefKataban)) = "UCAC2-TS-" & _
                            '                                               objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                            '    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(10).Trim)
                            'End If
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "UCAC2" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
