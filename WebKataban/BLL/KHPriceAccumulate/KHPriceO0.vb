'************************************************************************************
'*  ProgramID  ：KHPriceO0
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/27   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：高剛性タイプ　リニアスライドシリンダ　ＬＣＧ／ＬＣＲ／ＬＣＸ
'*
'*  更新履歴   ：                        更新日：2008/04/08   更新者：T.Sato
'*  受付No.RM0803014対応  オプション４について『U』と『P7*』を複数選択できるように修正
'*                        それに伴い価格計算ロジックも複数選択に対応
'*                                      更新日：2009/05/26   更新者：M.Miura
'*  ・受付No：RM0906034  二次電池対応機器　LCG
'*                                      更新日：2009/08/18   更新者：Y.Miura
'*  ・受付No：RM0906034  新機種追加　LCX
'*                                      更新日：2009/09/03   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'************************************************************************************
Module KHPriceO0

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer = 0
        Dim bolC5Flag As Boolean
        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim bolOptionP4 As Boolean = False              'RM0906034 2009/08/18 Y.Miura　二次電池対応
        Dim strOptionP4 As String = String.Empty        'RM0906034 2009/08/28 Y.Miura　二次電池対応

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban, 1, 3)
                '2012/02/17 変更
                Case "LCG"
                    'オプション加算価格キー
                    'RM0906034 2009/08/18 Y.Miura　二次電池対応
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(12), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "P4", "P40"
                                bolOptionP4 = True
                                strOptionP4 = strOpArray(intLoopCnt).Trim
                        End Select
                    Next
                    'RM1712027_LCV追加
                Case "LCV"
                Case Else
                    'オプション加算価格キー
                    'RM0906034 2009/08/18 Y.Miura　二次電池対応
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "P4", "P40"
                                bolOptionP4 = True
                                strOptionP4 = strOpArray(intLoopCnt).Trim
                        End Select
                    Next
            End Select
            'C5チェック
            'RM0906034 2009/08/18 Y.Miura　二次電池対応
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc)
            'bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

            'ストローク設定
            intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(1).Trim), _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim))

            '基本価格キー
            Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban, 1, 3)
                Case "LCX"
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim.Length <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                   intStroke.ToString
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   intStroke.ToString
                    End If

                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                    'RM1712027_LCV追加
                Case "LCV"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strKeyKataban & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1

                    'RM1807055_二次電池機種追加
                    If objKtbnStrc.strcSelection.strKeyKataban = "P" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strKeyKataban & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   intStroke.ToString & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
            End Select

            'バリエーション加算価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                'Case "LCG-Q"
                Case "LCG-Q", "LCR-Q", "LCX-Q"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
            End Select

            'RM1712027_LCV追加
            If Mid(objKtbnStrc.strcSelection.strSeriesKataban, 1, 3) = "LCV" Then
                'スイッチ加算価格キー
                If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strKeyKataban & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)

                    'RM1807055_二次電池機種追加の為修正
                    If objKtbnStrc.strcSelection.strKeyKataban = "P" Then
                    Else
                        'リード線長さ加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strKeyKataban & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                        End If

                    End If
                End If

                'オプション(1)加算価格キー
                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strKeyKataban & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strKeyKataban & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    End If

                    decOpAmount(UBound(decOpAmount)) = 1
                End If

                ''オプション(3)加算価格キー
                'If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                '    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                '                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim & CdCst.Sign.Hypen & _
                '                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                '    decOpAmount(UBound(decOpAmount)) = 1
                '    If bolC5Flag = True Then
                '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                '    End If
                'End If
            Else
                'スイッチ加算価格キー
                If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算不要
                    'If bolC5Flag = True Then
                    '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    'End If

                    'リード線長さ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                        'If bolC5Flag = True Then
                        '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        'End If
                    End If

                    'RM0906034 2009/08/18 Y.Miura　二次電池対応
                    If bolOptionP4 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-P4"
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                    End If
                End If


                'オプション(1)加算価格キー
                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    Else
                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        End If
                    End If

                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                    'RM0906034 2009/08/28 Y.Miura　二次電池対応
                    If bolOptionP4 Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            Case "A1", "A2", "A3", "A4" 'ショックキラー付
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                          "A" & CdCst.Sign.Hypen & _
                                                                          strOptionP4 & CdCst.Sign.Hypen & _
                                                                          objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "A5", "A6"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                          "A" & CdCst.Sign.Hypen & _
                                                                          strOptionP4 & CdCst.Sign.Hypen & _
                                                                          objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 2
                        End Select
                    End If
                End If
                'Else
                'ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                'ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                '                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                '                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                'decOpAmount(UBound(decOpAmount)) = 1
                'If bolC5Flag = True Then
                '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                'End If
                'End If
                'End If

                'オプション(3)加算価格キー
                If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If
            End If

            '2012/02/17 変更
            Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban, 1, 3)
                Case "LCG"
                    'クリーン仕様加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select
                    Next
                    'クリーン仕様加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(11), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "U"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'RM0906034 2009/08/18 Y.Miura　二次電池対応
                                'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                '                                           Left(strOpArray(intLoopCnt).Trim, 2) & CdCst.Sign.Hypen & _
                                '                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select
                    Next
                    'クリーン仕様加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(12), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                                'RM0906034 2009/08/18 Y.Miura　二次電池対応
                            Case "P4", "P40"    '二次電池
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算不要
                                'If bolC5Flag = True Then
                                '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                'End If
                            Case "P72", "P73"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                           Left(strOpArray(intLoopCnt).Trim, 2) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算不要
                                'If bolC5Flag = True Then
                                '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                'End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select
                    Next
                    'RM1712027_LCV追加
                Case "LCV"
                Case Else
                    'クリーン仕様加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "U"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                                'RM0906034 2009/08/18 Y.Miura　二次電池対応
                            Case "P4", "P40"    '二次電池
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算不要
                                'If bolC5Flag = True Then
                                '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                'End If
                            Case "P72", "P73"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                           Left(strOpArray(intLoopCnt).Trim, 2) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算不要
                                'If bolC5Flag = True Then
                                '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                'End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'RM0906034 2009/08/18 Y.Miura　二次電池対応
                                'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                '                                           Left(strOpArray(intLoopCnt).Trim, 2) & CdCst.Sign.Hypen & _
                                '                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select
                    Next
            End Select

            'オプション(5)加算価格キー(LCXシリーズ　キー形番"1"のみ）
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "LCX"
                    If objKtbnStrc.strcSelection.strKeyKataban = "1" Then
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(11), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case ""
                                Case "P72", "P73"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                               Left(strOpArray(intLoopCnt).Trim, 2) & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1

                            End Select
                        Next
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
