'************************************************************************************
'*  ProgramID  ：KHPrice75
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＡＸ１０００
'*  　　       ：ＡＸ２０００
'*  　　       ：ＡＸ３０００
'*  　　       ：ＡＸ４０００（Ｇ）
'*  　　       ：ＡＸ５０００
'*  　　       ：ＡＸ８０００
'*  　　       ：ＡＸ６０００
'*
'*  ・受付No：RM0907072  新型アブソデックス追加（AX1000T/AX2000T/AX4000T）
'*                                      更新日：2009/08/17   更新者：Y.Miura
'*  ・受付No：RM0908025  インターフェース仕様にCC-Linkを追加（AX1000T/AX2000T/AX4000T）
'*                                      更新日：2009/09/02   更新者：Y.Miura
'************************************************************************************
Module KHPrice75

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            '↓RM1310004 2013/10/01 追加
            Select Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3)
                Case "AX6"
                    '取付ベース加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim.Length <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    'ケーブル長さ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim.Length <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    'インターフェース仕様
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "M" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "U0" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                    '201503月次更新
                Case "AX7"
                    '取付ベース加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim.Length <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "X" Then
                        'ケーブル長さ加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim.Length <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        'ドライバ電源電圧
                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim.Length <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    End If
                Case Else
                    '中空固定軸加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0907072 2009/08/17 Y.Miura
                        'Select Case True
                        '    Case (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "006") Or _
                        '         (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "012")
                        '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                        '    Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX3"
                        '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                        'End Select
                    End If

                    '取付ベース加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0907072 2009/08/17 Y.Miura
                        'Select Case True
                        '    Case (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "006") Or _
                        '         (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "012")
                        '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                        '    Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX3"
                        '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                        'End Select
                    End If

                    'コネクタ取付方向加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0907072 2009/08/17 Y.Miura
                        'Select Case True
                        '    Case (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "006") Or _
                        '         (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "012")
                        '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                        '    Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX3"
                        '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                        'End Select
                    End If

                    'ケーブル変更加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "K" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            'RM0907072 2009/08/17 Y.Miura
                            'Select Case True
                            '    Case (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "006") Or _
                            '         (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "012")
                            '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                            '    Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX3"
                            '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                            'End Select
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            'RM0907072 2009/08/17 Y.Miura
                            'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***" & CdCst.Sign.Hypen & _
                            '                                           objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                Case "TS", "TH"
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***T" & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                Case Else
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***" & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            End Select
                            decOpAmount(UBound(decOpAmount)) = 1
                            'RM0907072 2009/08/17 Y.Miura
                            'Select Case True
                            '    Case (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "006") Or _
                            '         (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "012")
                            '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                            '    Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX3"
                            '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                            'End Select
                        End If
                    End If

                    'ブレーキ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0907072 2009/08/17 Y.Miura
                        'Select Case True
                        '    Case (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "006") Or _
                        '         (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "012")
                        '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                        '    Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX3"
                        '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                        'End Select
                    End If

                    'ドライバ電源電圧加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                'RM0907072 2009/08/17 Y.Miura
                                'Select Case True
                                '    Case (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "006") Or _
                                '         (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "012")
                                '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                                '    Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX3"
                                '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                                'End Select
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                'RM0907072 2009/08/17 Y.Miura
                                'Select Case True
                                '    Case (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "006") Or _
                                '         (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "012")
                                '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                                '    Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX3"
                                '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                                'End Select
                            End If
                        Else
                            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                'RM0907072 2009/08/17 Y.Miura
                                'Select Case True
                                '    Case (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "006") Or _
                                '         (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "012")
                                '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                                '    Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX3"
                                '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                                'End Select
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "***"
                                decOpAmount(UBound(decOpAmount)) = 1
                                'RM0907072 2009/08/17 Y.Miura
                                'Select Case True
                                '    Case (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "006") Or _
                                '         (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "012")
                                '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                                '    Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX3"
                                '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                                'End Select
                            End If
                        End If
                    End If

                    'ノックピン加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(10).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0907072 2009/08/17 Y.Miura
                        'Select Case True
                        '    Case (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "006") Or _
                        '         (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "012")
                        '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                        '    Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX3"
                        '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                        'End Select
                    End If

                    '本体表面処理加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(11).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0907072 2009/08/17 Y.Miura
                        'Select Case True
                        '    Case (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "006") Or _
                        '         (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "012")
                        '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                        '    Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX3"
                        '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                        'End Select
                    End If

                    'テーブル形状加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(12).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0907072 2009/08/17 Y.Miura
                        'Select Case True
                        '    Case (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "006") Or _
                        '         (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX2" And objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "012")
                        '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                        '    Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) = "AX3"
                        '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Open
                        'End Select
                    End If

                    'RM0908025 2009/09/02 Y.Miura
                    'インターフェース仕様加算価格キー
                    '2009/10/06 Y.Miura 不具合対応　13番目の要素はキー形番「T」の時だけ存在する
                    If objKtbnStrc.strcSelection.strOpSymbol.Length > 13 Then       'RM0912039 オプション追加による表示不具合修正
                        If objKtbnStrc.strcSelection.strKeyKataban = "T" Then
                            If objKtbnStrc.strcSelection.strOpSymbol(13).Trim <> "" Then

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                           "***" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
