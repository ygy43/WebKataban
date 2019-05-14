'************************************************************************************
'*  ProgramID  ：KHPriceD6
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/07   作成者：NII K.Sudoh
'*
'*  概要       ：ＦＲＬ白色シリーズ
'*
'*【更新履歴】
'*                                      更新日：2007/10/17   更新者：NII A.Takahashi
'*  ・クリーン仕様追加により修正
'*                                      更新日：2008/03/24   更新者：NII A.Takahashi
'*  ・ねじ・アタッチメント追加により修正
'*                                      更新日：2008/09/12   更新者：T.Sato
'*  ・RM0808096対応　ＭＸ白色シリーズ　Ｇネジ、ＮＰＴネジ、表示単位追加
'*  ・受付No：RM0904032  FRL2000新発売
'*                                      更新日：2009/06/18   更新者：Y.Miura
'*  ・受付No：RM0907070  二次電池対応機器　F3000/F4000/F6000/M3000/M4000
'*                                      更新日：2009/08/24   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　KHOptionCtl.vb
'*                                      更新日：2010/02/22   更新者：Y.Miura
'************************************************************************************
Module KHPriceD6

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim bolOptionF As Boolean = False
        Dim bolOptionFF As Boolean = False
        Dim bolOptionY As Boolean = False
        Dim bolOptionT As Boolean = False
        Dim bolOptionS As Boolean = False
        Dim bolOptionX As Boolean = False
        Dim bolOptionR As Boolean = False
        Dim bolOptionF1 As Boolean = False      'RM0904032 2009/06/18 Y.Miura
        Dim bolOptionP4 As Boolean = False      'RM0907070 2009/08/24 Y.Miura　二次電池対応

        Dim intOptionPos As Integer
        Dim strOptionP7 As String

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            '初期化
            strOptionP7 = ""

            If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 5 Then
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "P74" Then
                    strOptionP7 = "P74"
                End If
            End If
            If strOptionP7 = "P74" Then
                '基本価格キー
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                           "W" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                '配管アダプタセットアタッチメント価格キー
                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    Case "R2000", "R2100"
                        intOptionPos = 6
                    Case Else
                        intOptionPos = 4
                End Select

                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       "W" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Next
            Else
                'オプション選択判定
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        'RM0904032 2009/06/18 Y.Miura
                        'Case "F", "F1"
                        '    bolOptionF = True
                        Case "F"
                            bolOptionF = True
                        Case "F1"
                            bolOptionF = True
                            bolOptionF1 = True
                        Case "FF", "FF1"
                            bolOptionFF = True
                        Case "Y"
                            bolOptionY = True
                        Case "T"
                            bolOptionT = True
                        Case "S"
                            bolOptionS = True
                        Case "X"
                            bolOptionX = True
                            'Case "R1"
                        Case "R1", "RN", "RP" 'RM1610009
                            bolOptionR = True
                            'RM0907070 2009/08/24 Y.Miura　二次電池対応
                            'Case "P4", "P40"
                            '    bolOptionP4 = True
                    End Select
                Next

                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    'RM0904032 2009/06/18 Y.Miura
                    'Case "C3000", "C4000", "C1000", "C8000", "C2500", "C6500"
                    Case "C1000", "C2000", "C2500", "C3000", "C4000", "C8000", "C6500"
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim


                        If bolOptionF = True Then

                            'RM0904032 2009/06/18 Y.Miura
                            'strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "F"
                            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "C2000" And bolOptionF1 Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "F1"
                            Else
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "F"
                            End If

                            If bolOptionY = True Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"

                                If bolOptionT = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                                End If
                            Else
                                If bolOptionT = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                                End If
                            End If
                        Else
                            If bolOptionFF = True Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "FF"

                                If bolOptionY = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                                End If
                                If bolOptionT = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                                End If
                            Else
                                If bolOptionY = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "Y"

                                    If bolOptionT = True Then
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                                    End If
                                Else
                                    If bolOptionT = True Then
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "T"
                                    End If
                                End If
                            End If
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0904032 2009/06/18 Y.Miura
                        'Case "C3010", "C3020", "C3030", "C3040", _
                        '     "C3050", "C3060", "C3070", _
                        '     "C4010", "C4020", "C4030", "C4040", _
                        '     "C4050", "C4060", "C4070", _
                        '     "C1010", "C1020", "C1030", "C1040", _
                        '     "C1050", "C1060", _
                        '     "C8010", "C8020", "C8030", "C8040", _
                        '     "C8050", "C8060", "C8070", _
                        '     "C2520", "C2530", "C2550", _
                        '     "C6020", "C6030", "C6050", "C6060", "C6070"
                    Case "C1010", "C1020", "C1030", "C1040", "C1050", "C1060", _
                         "C2010", "C2020", "C2030", "C2040", "C2050", "C2060", _
                         "C2520", "C2530", "C2550", _
                         "C3010", "C3020", "C3030", "C3040", "C3050", "C3060", "C3070", _
                         "C4010", "C4020", "C4030", "C4040", "C4050", "C4060", "C4070", _
                         "C8010", "C8020", "C8030", "C8040", "C8050", "C8060", "C8070", _
                         "C6020", "C6030", "C6050", "C6060", "C6070"
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0904032 2009/06/18 Y.Miura
                        'Case "W3000", "W3100", "W4000", "W4100", _
                        '     "W1000", "W1100", "W8000", "W8100"

                        Select Case objKtbnStrc.strcSelection.strKeyKataban
                            Case "X"
                                '食品製造工程向け商品
                                If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "W" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select

                    Case "W1000", "W1100", "W2000", "W2100", "W3000", "W3100", _
                         "W4000", "W4100", "W8000", "W8100"
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim

                        If bolOptionF = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "F"

                            If bolOptionY = True Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"

                                If bolOptionT = True Or bolOptionR = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                                End If
                            Else
                                If bolOptionT = True Or bolOptionR = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                                End If
                            End If
                        Else
                            If bolOptionFF = True Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "FF"

                                If bolOptionY = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                                End If
                                If bolOptionT = True Or bolOptionR = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                                End If
                            Else
                                If bolOptionY = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "Y"

                                    If bolOptionT = True Or bolOptionR = True Then
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                                    End If
                                Else
                                    If bolOptionT = True Or bolOptionR = True Then
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "T"
                                    End If
                                End If
                            End If
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1

                        Select Case objKtbnStrc.strcSelection.strKeyKataban
                            Case "F"
                                '食品製造工程向け商品
                                If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "W" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If

                        End Select

                    Case "FW4000", "FW8000", "WW4000", "WW8000", "RW8000", "RW4000", _
                     "MW4000", "MW8000", "LW4000", "LW8000"
                        'RM1402017 フィルタレギュレータ　機種追加
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim

                        If bolOptionF1 = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "F1"

                            If bolOptionY = True Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                            End If

                            If bolOptionS = True Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "S"
                            End If

                        Else
                            If bolOptionF = True Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "F"

                                If bolOptionY = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                                End If
                                If bolOptionS = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "S"
                                End If
                            Else
                                If bolOptionY = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "Y"
                                End If
                                If bolOptionS = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "S"
                                End If
                            End If
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1

                        'RM0904032 2009/06/18 Y.Miura
                        'Case "M3000", "M4000", "M1000", "M8000", "M6000"
                    Case "M1000", "M2000", "M3000", "M4000", "M6000", "M8000"
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim

                        If bolOptionF = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "F1"

                            If bolOptionS = True Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "S"
                            End If
                        Else
                            If bolOptionS = True Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "S"
                            End If
                        End If

                        If bolOptionX = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "X"
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0904032 2009/06/18 Y.Miura
                        'Case "F3000", "F4000", "F1000", "F8000", "F6000"

                        Select Case objKtbnStrc.strcSelection.strKeyKataban
                            Case "F"
                                '食品製造工程向け商品
                                If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "W" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If

                        End Select

                    Case "F1000", "F2000", "F3000", "F4000", "F6000", "F8000"
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim

                        If bolOptionF = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "F"

                            If bolOptionY = True Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                            End If
                        Else
                            If bolOptionFF = True Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "FF"

                                If bolOptionY = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                                End If
                            Else
                                If bolOptionY = True Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "Y"
                                End If
                            End If
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1

                        Select Case objKtbnStrc.strcSelection.strKeyKataban
                            Case "F"
                                '食品製造工程向け商品
                                If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "W" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If

                        End Select

                    Case "MX3000", "MX4000", "MX1000", "MX8000", "MX6000"
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim

                        If bolOptionF = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "F1"
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1

                        Select Case objKtbnStrc.strcSelection.strKeyKataban
                            Case "F"
                                '食品製造工程向け商品
                                If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "W" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If

                            Case Else

                                '表示単位
                                If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "J1"
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select

                    Case "R3000", "R3100", "R4000", "R4100", _
                         "R1000", "R1100", "R8000", "R8100", _
                         "R6000", "R6100"
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim

                        If bolOptionT = True Or bolOptionR = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "T"
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1

                        Select Case objKtbnStrc.strcSelection.strKeyKataban
                            Case "F"
                                '食品製造工程向け商品
                                If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "W" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select

                    Case "R2000", "R2100"
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   "BASE" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim

                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "J1" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "J1"
                        End If

                        If bolOptionT = True Or bolOptionR = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "T"
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(5).Trim.Length <> 0 Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        End If

                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1

                        Select Case objKtbnStrc.strcSelection.strKeyKataban
                            Case "F"
                                '食品製造工程向け商品
                                If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "W" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select

                    Case "L3000", "L4000", "L1000", "L8000"
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select

                'RM1311028 2013/11/13 追加
                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    Case "W3000", "W3100"
                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "C" Then
                            'オプション価格加算キー
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "15W" & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                        Else
                            'オプション価格加算キー
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "W" & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                        End If
                    Case Else
                        'オプション価格加算キー
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case ""
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "W" & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Next
                End Select

                'RM1311028 2013/11/13 追加

                '組付けアタッチメント価格キー
                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    'RM0904032 2009/06/18 Y.Miura
                    'Case "C1000", "C1010", "C1020", "C1030", "C1040", "C1050", "C1060", _
                    '     "C2500", "C2520", "C2530", "C2550", _
                    '     "C3000", "C3010", "C3020", "C3030", "C3040", "C3050", "C3060", "C3070", _
                    '     "C4000", "C4010", "C4020", "C4030", "C4040", "C4050", "C4060", "C4070", _
                    '     "C6500", "C6020", "C6030", "C6050", "C6060", "C6070", _
                    '     "C8000", "C8010", "C8020", "C8030", "C8040", "C8050", "C8060", "C8070"
                    Case "C1000", "C1010", "C1020", "C1030", "C1040", "C1050", "C1060", _
                         "C2000", "C2010", "C2020", "C2030", "C2040", "C2050", "C2060", _
                         "C2500", "C2520", "C2530", "C2550", _
                         "C3000", "C3010", "C3020", "C3030", "C3040", "C3050", "C3060", "C3070", _
                         "C4000", "C4010", "C4020", "C4030", "C4040", "C4050", "C4060", "C4070", _
                         "C6500", "C6020", "C6030", "C6050", "C6060", "C6070", _
                         "C8000", "C8010", "C8020", "C8030", "C8040", "C8050", "C8060", "C8070", _
                         "W1000", "W2000", "W4000", "W1100", "W2100", "W4100"

                        If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                Case "C4020", "C4030"
                                    If (objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "20" Or _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "20N" Or _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "20G") And _
                                                                   (InStr(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, "S") <> 0) Then
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   "W" & CdCst.Sign.Hypen & "20" & CdCst.Sign.Hypen
                                    Else
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   "W" & CdCst.Sign.Hypen
                                    End If
                                Case Else
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "W" & CdCst.Sign.Hypen
                            End Select
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(5), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case Else
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)).Trim & strOpArray(intLoopCnt).Trim
                                End Select
                            Next
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case "W3000", "W3100"
                        If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                Case "W"
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "W" & CdCst.Sign.Hypen
                                Case "C"
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               "15W" & CdCst.Sign.Hypen
                            End Select
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(5), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case Else
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)).Trim & strOpArray(intLoopCnt).Trim
                                End Select
                            Next
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                End Select

                '配管アダプタセット・アタッチメント価格加算キー
                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    Case "R2000", "R2100"
                        intOptionPos = 6
                        'RM0904032 2009/06/18 Y.Miura
                        'Case "C1000", "C1010", "C1020", "C1030", "C1040", _
                        '     "C1050", "C1060", _
                        '     "C2500", "C2520", "C2530", "C2550", _
                        '     "C3000", "C3010", "C3020", "C3030", "C3040", _
                        '     "C3050", "C3060", "C3070", _
                        '     "C4000", "C4010", "C4020", "C4030", "C4040", _
                        '     "C4050", "C4060", "C4070", _
                        '     "C6500", "C6020", "C6030", "C6050", "C6060", "C6070", _
                        '     "C8000", "C8010", "C8020", "C8030", "C8040", _
                        '     "C8050", "C8060", "C8070"
                    Case "C1000", "C1010", "C1020", "C1030", "C1040", "C1050", "C1060", _
                         "C2000", "C2010", "C2020", "C2030", "C2040", "C2050", "C2060", _
                         "C2500", "C2520", "C2530", "C2550", _
                         "C3000", "C3010", "C3020", "C3030", "C3040", "C3050", "C3060", "C3070", _
                         "C4000", "C4010", "C4020", "C4030", "C4040", "C4050", "C4060", "C4070", _
                         "C6500", "C6020", "C6030", "C6050", "C6060", "C6070", _
                         "C8000", "C8010", "C8020", "C8030", "C8040", "C8050", "C8060", "C8070", _
                         "W1000", "W2000", "W3000", "W4000", "W1100", "W2100", "W3100", "W4100"
                        intOptionPos = 7
                    Case "MX1000", "MX3000", "MX4000", "MX6000", "MX8000"
                        intOptionPos = 5
                    Case "FW4000", "FW8000", "WW4000", "WW8000", "RW8000", "RW4000", _
                         "MW4000", "MW8000", "LW4000", "LW8000"
                        'RM1402017 機種追加
                        intOptionPos = 4
                    Case Else
                        intOptionPos = 5
                End Select

                'RM1311028 2013/11/13 追加
                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    Case "W3000", "W3100"
                        If objKtbnStrc.strcSelection.strKeyKataban.Trim = "C" Then
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "15W" & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                        Else
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "W" & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                        End If
                    Case Else
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case ""
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "W" & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Next
                End Select
            End If
        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
