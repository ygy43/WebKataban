'************************************************************************************
'*  ProgramID  ：KHPriceQ8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2009/08/11   作成者：Y.Miura
'*                                      更新日：             更新者：
'*
'*  概要       ：ESSD,ELCRシリーズ  (電動アクチュエータ)
'*
'************************************************************************************
Module KHPriceQ8

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            'RM1312084 2013/12/25
            'RM1402099 2014/02/25 ETSシリーズ追加
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim

                Case "ETV"
                    'RM1410045
                    If objKtbnStrc.strcSelection.strKeyKataban = "T" Or objKtbnStrc.strcSelection.strKeyKataban = "U" Or _
                        objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                        'TOYO品
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   "T" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '原点センサ
                        If objKtbnStrc.strcSelection.strOpSymbol(8) <> "N" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                        "T" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(10) <> "N" And _
                            objKtbnStrc.strcSelection.strOpSymbol(10) <> "D" Then
                            'グリースニップル
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                        "T" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(11) <> "N" Then
                            '第２オプション追加  2017/03/22 追加
                            'ボディサイズが「05」「06」のときとそれ以外のときで条件分け
                            If objKtbnStrc.strcSelection.strOpSymbol(1) <> "05" And _
                               objKtbnStrc.strcSelection.strOpSymbol(1) <> "06" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-10"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-05"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                            '食品
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            'オプション追加により、一つずれるため修正  2017/03/22 修正
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(12)
                            'strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                            '                                           objKtbnStrc.strcSelection.strOpSymbol(11)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    Else
                        '日本品
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '原点センサ
                        If objKtbnStrc.strcSelection.strOpSymbol(8) <> "N" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(10) <> "N" And _
                            objKtbnStrc.strcSelection.strOpSymbol(10) <> "D" Then
                            'グリースニップル
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(11) <> "N" Then
                            '第２オプション追加  RM1702018  2017/02/13 追加
                            'ボディサイズが「05」「06」のときとそれ以外のときで条件分け
                            If objKtbnStrc.strcSelection.strOpSymbol(1) <> "05" And _
                               objKtbnStrc.strcSelection.strOpSymbol(1) <> "06" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-10"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-05"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                            '食品
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            'オプション追加により、一つずれるため修正  RM1702018  2017/02/13 修正
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(12)
                            'strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                            '                                           objKtbnStrc.strcSelection.strOpSymbol(11)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                Case "ECS"
                    If objKtbnStrc.strcSelection.strKeyKataban = "T" Or objKtbnStrc.strcSelection.strKeyKataban = "U" Or _
                        objKtbnStrc.strcSelection.strKeyKataban = "V" Or objKtbnStrc.strcSelection.strKeyKataban = "W" Or _
                        objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                        'TOYO品
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   "T" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3)
                        decOpAmount(UBound(decOpAmount)) = 1

                        'モータ取付方法
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   "T" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1)
                        decOpAmount(UBound(decOpAmount)) = 1

                        If objKtbnStrc.strcSelection.strOpSymbol(8) <> "N" Then
                            '原点センサ
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       "T" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(10) <> "N" Then
                            'グリースニップル
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       "T" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(11) <> "N" Then
                            '第２オプション追加  2017/03/22 追加
                            'ボディサイズが「05」「06」のときとそれ以外のときで条件分け
                            If objKtbnStrc.strcSelection.strOpSymbol(1) <> "05" And _
                               objKtbnStrc.strcSelection.strOpSymbol(1) <> "06" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-10"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-05"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "V" Or objKtbnStrc.strcSelection.strKeyKataban = "W" Then
                            'オプション追加により、一つずれるため修正  2017/03/22 修正
                            If objKtbnStrc.strcSelection.strOpSymbol(12) <> "N" Then
                                '防錆処理
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(12)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            'オプション追加により、一つずれるため修正  2017/03/22 修正
                            If objKtbnStrc.strcSelection.strOpSymbol(13) <> String.Empty Then
                                '二次電池
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(13)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                            '食品
                            'オプション追加により、一つずれるため修正  2017/03/22 修正
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(12)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    Else
                        '日本品
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3)
                        decOpAmount(UBound(decOpAmount)) = 1

                        If objKtbnStrc.strcSelection.strOpSymbol(4) <> "E" And _
                            objKtbnStrc.strcSelection.strOpSymbol(4) <> "B" Then
                            'モータ取付方法
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(8) <> "N" Then
                            '原点センサ
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(10) <> "N" Then
                            'グリースニップル
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(11) <> "N" Then
                            '第２オプション追加  RM1702018  2017/02/13 追加
                            'ボディサイズが「05」「06」のときとそれ以外のときで条件分け
                            If objKtbnStrc.strcSelection.strOpSymbol(1) <> "05" And _
                               objKtbnStrc.strcSelection.strOpSymbol(1) <> "06" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-10"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-05"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "4" Then
                            'オプション追加により、一つずれるため修正  RM1702018  2017/02/13 修正
                            If objKtbnStrc.strcSelection.strOpSymbol(12) <> "N" Then
                                '防錆処理
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(12)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            'オプション追加により、一つずれるため修正  RM1702018  2017/02/13 修正
                            If objKtbnStrc.strcSelection.strOpSymbol(13) <> String.Empty Then
                                '二次電池
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(13)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                            '食品
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            'オプション追加により、一つずれるため修正  RM1702018  2017/02/13 修正
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(12)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                Case "ETS"
                    'RM1402053
                    If objKtbnStrc.strcSelection.strKeyKataban = "T" Or objKtbnStrc.strcSelection.strKeyKataban = "U" Or _
                        objKtbnStrc.strcSelection.strKeyKataban = "V" Or objKtbnStrc.strcSelection.strKeyKataban = "W" Or _
                         objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4)
                        decOpAmount(UBound(decOpAmount)) = 1

                        'モータ取付方法とストロークで減算(ボディサイズ：13,14,17)
                        If objKtbnStrc.strcSelection.strOpSymbol(1) = "13" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(1) = "14" Then

                            If objKtbnStrc.strcSelection.strOpSymbol(4) = "D" Then

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(4) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3)
                                decOpAmount(UBound(decOpAmount)) = 1

                            End If

                        End If
                        If objKtbnStrc.strcSelection.strOpSymbol(1) = "17" Then

                            If objKtbnStrc.strcSelection.strOpSymbol(4) = "D" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(4) = "R" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(4) = "L" Then

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(4) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3)
                                decOpAmount(UBound(decOpAmount)) = 1

                            End If

                        End If

                        '原点センサ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8)
                        decOpAmount(UBound(decOpAmount)) = 1

                        'グリースニップル
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(10) & "-OP"
                        decOpAmount(UBound(decOpAmount)) = 1

                        '第２オプション追加  2017/03/22 追加
                        'ボディサイズが「05」「06」のときとそれ以外のときで条件分け
                        If objKtbnStrc.strcSelection.strOpSymbol(1) <> "05" And _
                           objKtbnStrc.strcSelection.strOpSymbol(1) <> "06" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(11) & "-10"
                            decOpAmount(UBound(decOpAmount)) = 1

                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(11) & "-05"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "V" Or objKtbnStrc.strcSelection.strKeyKataban = "W" Then
                            'オプション追加により、一つずれるため修正  2017/03/22 修正
                            If objKtbnStrc.strcSelection.strOpSymbol(12) <> "N" Then
                                '防錆処理
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(12)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            If objKtbnStrc.strcSelection.strOpSymbol(13) <> String.Empty Then
                                '二次電池
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(13)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                            '食品
                            'オプション追加により、一つずれるため修正  2017/03/22 修正
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(12)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    ElseIf objKtbnStrc.strcSelection.strKeyKataban = "A" Or objKtbnStrc.strcSelection.strKeyKataban = "B" Or _
                           objKtbnStrc.strcSelection.strKeyKataban = "C" Or objKtbnStrc.strcSelection.strKeyKataban = "D" Then
                        '日本Multi Axisシリーズ
                        '基本価格
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & Left(objKtbnStrc.strcSelection.strOpSymbol(2), 1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3) & objKtbnStrc.strcSelection.strOpSymbol(4)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '原点・リミットセンサ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7)
                        decOpAmount(UBound(decOpAmount)) = 1

                    ElseIf objKtbnStrc.strcSelection.strKeyKataban = "I" Or objKtbnStrc.strcSelection.strKeyKataban = "J" Or _
                     objKtbnStrc.strcSelection.strKeyKataban = "K" Or objKtbnStrc.strcSelection.strKeyKataban = "L" Or _
                     objKtbnStrc.strcSelection.strKeyKataban = "M" Or objKtbnStrc.strcSelection.strKeyKataban = "N" Or _
                     objKtbnStrc.strcSelection.strKeyKataban = "O" Or objKtbnStrc.strcSelection.strKeyKataban = "P" Then
                        '日本Multi Axisシリーズ
                        '基本価格
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & Left(objKtbnStrc.strcSelection.strOpSymbol(2), 1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3) & objKtbnStrc.strcSelection.strOpSymbol(4)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '原点・リミットセンサ
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "210" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7) & CdCst.Sign.Hypen & "210"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    Else
                        '日本標準品
                        '基本価格
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3)
                        decOpAmount(UBound(decOpAmount)) = 1

                        'モータ取付方法
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '原点センサ
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8)
                        decOpAmount(UBound(decOpAmount)) = 1

                        'グリースニップル
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(10) & "-OP"
                        decOpAmount(UBound(decOpAmount)) = 1

                        '第２オプション追加  RM1702018  2017/02/13 追加
                        'ボディサイズが「05」「06」のときとそれ以外のときで条件分け
                        If objKtbnStrc.strcSelection.strOpSymbol(1) <> "05" And _
                           objKtbnStrc.strcSelection.strOpSymbol(1) <> "06" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(11) & "-10"
                            decOpAmount(UBound(decOpAmount)) = 1

                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(11) & "-05"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "4" Then
                            'オプション追加により、一つずれるため修正  RM1702018  2017/02/13 修正
                            If objKtbnStrc.strcSelection.strOpSymbol(12) <> "N" Then
                                '防錆処理
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(12)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            'オプション追加により、一つずれるため修正  RM1702018  2017/02/13 修正
                            If objKtbnStrc.strcSelection.strOpSymbol(13) <> String.Empty Then
                                '二次電池
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(13)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                            '食品
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            'オプション追加により、一つずれるため修正  RM1702018  2017/02/13 修正
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(12)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        'RM1802016  メタルストパ仕様追加
                        'ボディサイズが「12」のときとそれ以外のときで条件分け
                        If objKtbnStrc.strcSelection.strKeyKataban = "Z" Then
                            If objKtbnStrc.strcSelection.strOpSymbol(1) = "12" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(12) & "-12"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(12) & "-06"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If
                    End If

                Case "ECV"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3)
                    decOpAmount(UBound(decOpAmount)) = 1


                    '原点センサ
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(8)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'グリースニップル
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(10)
                    decOpAmount(UBound(decOpAmount)) = 1


                    '対象外となっているキー型番についても対象となるため修正  2017/03/22 修正
                    'If objKtbnStrc.strcSelection.strKeyKataban = "T" Or objKtbnStrc.strcSelection.strKeyKataban = "U" _
                    '   Or objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                    
                    '第２オプション追加  RM1702018  2017/02/13 追加
                    'ボディサイズが「05」「06」のときとそれ以外のときで条件分け
                    If objKtbnStrc.strcSelection.strOpSymbol(1) <> "05" And _
                           objKtbnStrc.strcSelection.strOpSymbol(1) <> "06" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(11) & "-10"
                        decOpAmount(UBound(decOpAmount)) = 1

                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(11) & "-05"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    
                    If objKtbnStrc.strcSelection.strKeyKataban = "F" Or _
                         objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                        '食品
                        'オプション追加により、一つずれるため修正  RM1702018  2017/02/13 修正
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(12)
                        'strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                        '                                           objKtbnStrc.strcSelection.strOpSymbol(11)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "ESM"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1)
                        Case "HDU", "TTU", "CA", "SE", "PP1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "VC"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5)
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "ST"
                            Dim intST As Integer = 0
                            intST = Math.Ceiling(objKtbnStrc.strcSelection.strOpSymbol(2) * 0.01) * 100
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                       intST
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "B"
                            Dim intST As Integer = 0
                            intST = Math.Ceiling(objKtbnStrc.strcSelection.strOpSymbol(3) * 0.01) * 100
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1)
                            decOpAmount(UBound(decOpAmount)) = intST
                    End Select

                Case "ERL"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & "S-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "A" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(9)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "ERL2"
                    '基本加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & "E-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'モータ取付方向加算キー
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "E" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                             objKtbnStrc.strcSelection.strOpSymbol(1) & objKtbnStrc.strcSelection.strOpSymbol(2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'ブレーキ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'オプション加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "N0" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "N" Then
                    Else

                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                            Case "A", "B"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9) & "-EC07"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Case "C", "D"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9) & "-EC63"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Case "E", "F"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9) & "-ECPT"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select

                    End If

                Case "ESD"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & "S-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "A" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(9)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(10)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "ESD2"
                    '基本加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & "E-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'モータ取付方向加算キー
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "E" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                             objKtbnStrc.strcSelection.strOpSymbol(1) & objKtbnStrc.strcSelection.strOpSymbol(2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'ブレーキ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'オプション加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "N0" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "N" Then
                    Else

                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                            Case "A", "B"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9) & "-EC07"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Case "C", "D"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9) & "-EC63"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "E", "F"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9) & "-ECPT"
                                decOpAmount(UBound(decOpAmount)) = 1

                        End Select

                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(10) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'RM1803042_EBS・EBR追加
                Case "EBS"

                    '基本価格加算キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(4) & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(8)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'センサ取付
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(9)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'センサ取付
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(10)
                    decOpAmount(UBound(decOpAmount)) = 1

                Case "EBR"

                    '基本価格加算キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(5) & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(9)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'センサ取付
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(10)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'センサ取付
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(11)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'RM1804032_EKS追加
                Case "EKS"

                    '基本価格加算キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(2) & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(5)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'モータ取付方法
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(8) & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(10)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'センサ取付
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "005" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(5) & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(11)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(11)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "BASE" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module

