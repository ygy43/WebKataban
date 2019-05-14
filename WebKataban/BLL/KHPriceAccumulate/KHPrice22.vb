'************************************************************************************
'*  ProgramID  ：KHPrice22
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/28   作成者：NII K.Sudoh
'*                                      更新日：2013/02/04   更新者：Y.Tachi
'*
'*  概要       ：セレックスシリンダ　ＳＣＳ
'*                                   ＳＣＳ２
'************************************************************************************
Module KHPrice22

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStrokeS1 As Integer = 0
        Dim intStrokeS2 As Integer = 0
        Dim bolC5Flag As Boolean

        Dim decLength As Decimal
        Dim decWFLength As Decimal
        Dim strStdWFLength As String = String.Empty

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'C5チェック
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

            ' 基本価格キー
            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                Case "", "2", "F", "4"
                    'ストローク設定(S1)
                    intStrokeS1 = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                            CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                            CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim))

                    '基本価格キー
                    'Pを含む場合は両ロッド加算
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") < 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-BASE-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                   intStrokeS1.ToString
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-BASE-D-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                   intStrokeS1.ToString
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    End If

                    '一定以上ストロークの加算(二圧検定料)
                    '口径が160以上の場合、ストロークが一定以上ならば9000円を加算する
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "160" Then
                        '1948以上ならば、9000円を加算する(1965->1948 2008/5/27対応)
                        'S1
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) >= 1948 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-STRADD"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    ElseIf objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "180" Then
                        '1526以上ならば、9000円を加算する(1552->1526 2008/5/27対応)
                        'S1
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) >= 1526 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-STRADD"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    ElseIf objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "200" Then
                        '946以上ならば、9000円を加算する(1000->946 2008/5/27対応)
                        'S1
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) >= 946 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-STRADD"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    ElseIf objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "250" Then
                        '752以上ならば、9000円を加算する(805->752 2008/5/27対応)
                        'S1
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) >= 752 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-STRADD"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                Case "B"
                    'ストローク設定(S1)
                    intStrokeS1 = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                            CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                            CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim))
                    'ストローク設定(S2)
                    intStrokeS2 = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                            CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                            CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim))

                    '基本価格キー
                    'S1価格
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-BASE-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS1.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If

                    'S2価格
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-BASE-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If

                    '一定以上ストローク加算
                    '口径が160以上の場合、ストローク(S1+S2)が一定以上ならば9000円を加算する(対応を削除 2008/5/27対応)
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "180" Then
                        Dim Var_chk
                        Var_chk = objKtbnStrc.strcSelection.strOpSymbol(1).Trim Like "*B*"
                        If Var_chk Then
                            '1481以上ならば、9000円を加算する(1552->1481 2008/5/27対応)
                            If Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <> 0 Then
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) >= 1481 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-" & "STRADD"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            End If
                        End If
                    ElseIf objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "200" Then
                        '892以上ならば、9000円を加算する(1000->892 2008/5/27対応)
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) >= 892 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-STRADD"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    ElseIf objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "250" Then
                        '690以上ならば、9000円を加算する(805->690 2008/5/27対応)
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) >= 690 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-STRADD"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                Case "D", "G"
                    'ストローク設定(S1)
                    intStrokeS1 = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                            CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                            CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim))

                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-BASE-D-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS1.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If

                    '一定以上ストロークの加算(二圧検定料)
                    '口径が160以上の場合、ストロークが一定以上ならば9000円を加算する(廃止2008/5/27対応)
                    '口径が160以上の場合、ストロークが1552以上ならば9000円を加算する(廃止2008/5/27対応)
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "200" Then
                        '946以上ならば、9000円を加算する(1000->946 2008/5/27対応)
                        'S1
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) >= 946 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-STRADD"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    ElseIf objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "250" Then
                        '752以上ならば、9000円を加算する(805->752 2008/5/27対応)
                        'S1
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) >= 752 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-STRADD"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
            End Select

            'バリエーション「B」価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-B-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「G」価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G") >= 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G1") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-G-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "", "F"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "B"
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                    Case "D", "G"
                        decOpAmount(UBound(decOpAmount)) = 2
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「G1」価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G1") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G1-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "", "F"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "B"
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                    Case "D", "G"
                        decOpAmount(UBound(decOpAmount)) = 2
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「H」価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("H") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-H-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "", "F"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case "D", "G"
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「L」価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("L") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-L-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "", "F", "4"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case "D", "G"
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「N」価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("N") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-N-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "", "2", "F", "4"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case "D", "G"
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「P」価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-P-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「T」価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T") >= 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T1") < 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T2") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-T-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "", "F"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case "D", "G"
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「T1」価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T1") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-T1-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "", "F"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case "D", "G"
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「T2」価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T2") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-T2-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "", "F"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case "D", "G"
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「W」価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-W-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '支持形式価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "CB", "TC", "TA", "TB", "TF", "TD", "TE"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SUPPORT-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
            End Select

            '配管ねじ(S1)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                Case "N", "G"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SCREW-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Screw
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = strPriceDiv(UBound(strPriceDiv)) & CdCst.Sign.Delimiter.Pipe & CdCst.PriceAccDiv.C5
                    End If
            End Select

            '配管ねじ(S2)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                Case "N", "G"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SCREW-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(10).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Screw
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = strPriceDiv(UBound(strPriceDiv)) & CdCst.Sign.Delimiter.Pipe & CdCst.PriceAccDiv.C5
                    End If
            End Select

            'スイッチ(S1)
            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                'スイッチ加算価格キー
                Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                    Case "A", "B"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SW-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SW-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)

                        'リード線長さ加算価格キー
                        If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "SCS2" Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                Case "3", "5"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                        Case "T0H", "T0V", "T5H", "T5V", "T8H", "T8V", _
                                             "T1H", "T1V", "T2H", "T2V", "T3H", "T3V", _
                                             "T3PH", "T3PV", "T2WH", "T2WV", "T2YH", "T2YV", _
                                             "T3WH", "T3WV", "T3YH", "T3YV", "T2JH", "T2JV"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(1)-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                        Case "T2YDP", "T2YDUP"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(2)-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                        Case "T2YDPT"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(3)-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                    End Select
                            End Select
                        Else
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                Case "3", "5"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                        Case "R1K", "R2K", "R2YK", "R3K", "R3YK", _
                                             "R1KA", "R2KA", "R2YKA", "R3KA", "R3YKA", _
                                             "R1KB", "R2KB", "R2YKB", "R3KB", "R3YKB", _
                                             "R0", "R4", "R5", "R6", "R0B", _
                                             "R4B", "R5B", "R6B", "R0A", "R4A", _
                                             "R5A", "R6A"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(1)-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                        Case "T2YDP", "T2YDUP"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(2)-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                        Case "T2YDPT"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(3)-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                    End Select
                            End Select
                        End If
                End Select
            End If

            'スイッチ(S2)
            If objKtbnStrc.strcSelection.strOpSymbol(14).Trim <> "" Then
                'スイッチ加算価格キー
                Select Case objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                    Case "A", "B"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SW-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(14).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SW-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)

                        'リード線長さ加算価格キー
                        If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "SCS2" Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                                Case "3", "5"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                                        Case "T0H", "T0V", "T5H", "T5V", "T8H", "T8V", _
                                             "T1H", "T1V", "T2H", "T2V", "T3H", "T3V", _
                                             "T3PH", "T3PV", "T2WH", "T2WV", "T2YH", "T2YV", _
                                             "T3WH", "T3WV", "T3YH", "T3YV", "T2JH", "T2JV"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(1)-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                                        Case "T2YD"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(2)-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                                        Case "T2YDT"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(3)-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                                    End Select
                            End Select
                        Else
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                                Case "3", "5"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                                        Case "R1K", "R2K", "R2YK", "R3K", "R3YK", _
                                             "R1KA", "R2KA", "R2YKA", "R3KA", "R3YKA", _
                                             "R1KB", "R2KB", "R2YKB", "R3KB", "R3YKB", _
                                             "R0", "R4", "R5", "R6", "R0B", _
                                             "R4B", "R5B", "R6B", "R0A", "R4A", _
                                             "R5A", "R6A"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(1)-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                                        Case "T2YDP", "T2YDUP"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(2)-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                                        Case "T2YDPT"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(3)-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                                    End Select
                            End Select
                        End If
                End Select
            End If

            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(17), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "C2"
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "", "2", "F"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            Case "B"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "B", "R", "H"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                                            Case "B", "R", "H"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                decOpAmount(UBound(decOpAmount)) = 2
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                            Case Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                        End Select
                                    Case Else
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                                            Case "B", "R", "H"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                        End Select
                                End Select
                            Case "D", "G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select
                    Case "J", "K", "L"
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "", "2", "F"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS1.ToString
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            Case "B"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") < 0 Then
                                    'S1
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                               intStrokeS1.ToString
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                    End If

                                    'S2
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                               intStrokeS2.ToString
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                    End If
                                Else
                                    'S1
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                               intStrokeS1.ToString
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                    End If
                                End If
                            Case "D", "G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS1.ToString
                                decOpAmount(UBound(decOpAmount)) = 2
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select
                    Case "M"
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "", "2", "F", "4"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS1.ToString
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") < 0 Then
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 2
                                End If
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            Case "B"
                                'S1
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS1.ToString
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If

                                'S2
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS2.ToString
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            Case "D", "G", "H"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS1.ToString
                                decOpAmount(UBound(decOpAmount)) = 2
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select
                    Case "S", "T"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "", "2", "F"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "B"
                                decOpAmount(UBound(decOpAmount)) = 2
                            Case "D", "G"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case "P6"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "", "F"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("N") >= 0 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            Case "B"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 2
                                End If
                            Case "D", "G"
                                decOpAmount(UBound(decOpAmount)) = 2
                        End Select
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case "U1"
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "", "2", "F"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS1
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            Case "B"
                                'S1
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS1
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If

                                'S2
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS2
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            Case "D", "G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS1
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select
                    Case "A2"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "", "2", "F"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "B"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 2
                                End If
                            Case "D", "G"
                                If objKtbnStrc.strcSelection.strFullKataban.IndexOf("N13-N11") < 0 Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case "P4"
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "4"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 100 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-100-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 100 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 200 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-200-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 200 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 300 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-300-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 300 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 400 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-400-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 400 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 500 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-500-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 500 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 600 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-600-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 600 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 700 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-700-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 700 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 800 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-800-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 800 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 900 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-900-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 900 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 1000 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-1000-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 1000 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 1100 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-1100-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 1100 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 1200 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-1200-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 1200 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 1300 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-1300-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 1300 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 1400 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-1400-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 1400 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 1500 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-1500-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 1500 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 1600 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-1600-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 1600 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 1700 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-1700-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 1700 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 1800 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-1800-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 1800 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 1900 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-1900-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 1900 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 2000 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-2000-P4"
                                End If

                                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = 250 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 700 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-752-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = 200 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 900 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-945-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = 180 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 1500 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-1525-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = 160 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 1900 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-1947-P4"
                                End If
                                decOpAmount(UBound(decOpAmount)) = 1

                                'C5チェック
                                bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)
                        End Select
                End Select
            Next

            '付属品加算価格キー
            If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "SCS2" Then
                'SCS2
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(18), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case "S", "T", "R"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim

                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If

                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                Case "", "2", "F", "4"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "B"
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Case "D", "G"
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Case "P6"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                Case "", "F"
                                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("N") >= 0 And _
                                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Then
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                                Case "B"
                                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    End If
                                Case "D", "G"
                                    decOpAmount(UBound(decOpAmount)) = 2
                            End Select
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        Case "A2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                Case "", "2", "F"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "B"
                                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    End If
                                Case "D", "G"
                                    If objKtbnStrc.strcSelection.strFullKataban.IndexOf("N13-N11") < 0 Then
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                            End Select
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                    End Select
                Next
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(19), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case "I"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                Case "", "2", "F", "4"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "B"
                                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    End If
                                Case "D", "G"
                                    If objKtbnStrc.strcSelection.strFullKataban.IndexOf("N13-N11") < 0 Then
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                            End Select
                        Case "Y"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                Case "", "2", "F", "4"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "B"
                                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    End If
                                Case "D", "G"
                                    If objKtbnStrc.strcSelection.strFullKataban.IndexOf("N13-N11") < 0 Then
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                            End Select
                        Case "IY"
                            'I加算
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-I-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'Y加算
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-Y-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "B1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            'RM1307003 2013/07/04 追加
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                Case "SCS2"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "B"
                                            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(1), "B") = 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            End If
                                        Case "D", "G"
                                            If InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "N13-N11") = 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            End If
                                    End Select
                            End Select
                        Case "B2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            'RM1307003 2013/07/04 追加
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                Case "SCS2"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "B"
                                            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(1), "B") = 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            End If
                                        Case "D", "G"
                                            If InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "N13-N11") = 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            End If
                                    End Select
                            End Select
                        Case "FP1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                    End Select
                Next

                Select Case objKtbnStrc.strcSelection.strKeyKataban
                    Case "F", "G"
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(20), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case ""
                                Case "I"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "", "2", "F"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "B"
                                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            End If
                                        Case "D", "G"
                                            If objKtbnStrc.strcSelection.strFullKataban.IndexOf("N13-N11") < 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            End If
                                    End Select
                                Case "Y"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "", "2", "F"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "B"
                                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            End If
                                        Case "D", "G"
                                            If objKtbnStrc.strcSelection.strFullKataban.IndexOf("N13-N11") < 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            End If
                                    End Select
                                Case "IY"
                                    'I加算
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-I-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1

                                    'Y加算
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-Y-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "B1"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    'RM1307003 2013/07/04 追加
                                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                        Case "SCS2"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                Case "B"
                                                    If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(1), "B") = 0 Then
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    Else
                                                        decOpAmount(UBound(decOpAmount)) = 2
                                                    End If
                                                Case "D", "G"
                                                    If InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "N13-N11") = 0 Then
                                                        decOpAmount(UBound(decOpAmount)) = 2
                                                    Else
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    End If
                                            End Select
                                    End Select
                                Case "B2"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    'RM1307003 2013/07/04 追加
                                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                        Case "SCS2"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                Case "B"
                                                    If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(1), "B") = 0 Then
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    Else
                                                        decOpAmount(UBound(decOpAmount)) = 2
                                                    End If
                                                Case "D", "G"
                                                    If InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "N13-N11") = 0 Then
                                                        decOpAmount(UBound(decOpAmount)) = 2
                                                    Else
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    End If
                                            End Select
                                    End Select
                            End Select
                        Next
                    Case Else

                End Select

            Else
                'SCS
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "2"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(18).Trim
                            Case "P4"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 100 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-100-P4"
                                Else
                                    If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 100 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 200 Then
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-200-P4"
                                    Else
                                        If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 200 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 300 Then
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-300-P4"
                                        Else
                                            If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 300 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 400 Then
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-400-P4"
                                            Else
                                                If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 400 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 500 Then
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-500-P4"
                                                Else
                                                    If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 500 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 600 Then
                                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-600-P4"
                                                    Else
                                                        If objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 600 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 700 Then
                                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-700-P4"
                                                        Else
                                                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = 125 Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = 140 Or _
                                                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim = 160 Or objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 700 Then
                                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-800-P4"
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = 250 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 700 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-752-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = 180 Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = 200 And _
                                   objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 700 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 800 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-800-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = 180 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 800 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-900-P4"
                                End If
                                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = 200 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 800 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim <= 900 Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-900-P4"
                                Else
                                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = 200 And objKtbnStrc.strcSelection.strOpSymbol(12).Trim > 900 Then
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-945-P4"
                                    End If
                                End If
                                decOpAmount(UBound(decOpAmount)) = 1

                                'C5チェック
                                bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)
                                'RM1210067 2013/04/04 ローカル版との差異修正
                                'If bolC5Flag = True Then
                                '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                'End If
                        End Select

                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(19), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case ""
                                Case "I"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "", "2"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "B"
                                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            End If
                                        Case "D"
                                            If objKtbnStrc.strcSelection.strFullKataban.IndexOf("N13-N11") < 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            End If
                                    End Select
                                Case "Y"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "", "2"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "B"
                                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            End If
                                        Case "D"
                                            If objKtbnStrc.strcSelection.strFullKataban.IndexOf("N13-N11") < 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            End If
                                    End Select
                                Case "IY"
                                    'I加算
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-I-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1

                                    'Y加算
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-Y-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "B1"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "B2"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Next
                    Case Else
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(18), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case ""
                                Case "I"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "", "2"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "B"
                                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            End If
                                        Case "D"
                                            If objKtbnStrc.strcSelection.strFullKataban.IndexOf("N13-N11") < 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            End If
                                    End Select
                                Case "Y"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "", "2"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "B"
                                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            End If
                                        Case "D"
                                            If objKtbnStrc.strcSelection.strFullKataban.IndexOf("N13-N11") < 0 Then
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            Else
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            End If
                                    End Select
                                Case "IY"
                                    'I加算
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-I-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1

                                    'Y加算
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-Y-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "B1"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    'RM1307003 2013/07/04 追加
                                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                        Case "SCS2"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "SCS"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                Case "B"
                                                    If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(1), "B") = 0 Then
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    Else
                                                        decOpAmount(UBound(decOpAmount)) = 2
                                                    End If
                                                Case "D"
                                                    If InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "N13-N11") = 0 Then
                                                        decOpAmount(UBound(decOpAmount)) = 2
                                                    Else
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    End If
                                            End Select
                                    End Select
                                Case "B2"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    'RM1307003 2013/07/04 追加
                                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                        Case "SCS2"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "SCS"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                Case "B"
                                                    If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(1), "B") = 0 Then
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    Else
                                                        decOpAmount(UBound(decOpAmount)) = 2
                                                    End If
                                                Case "D"
                                                    If InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "N13-N11") = 0 Then
                                                        decOpAmount(UBound(decOpAmount)) = 2
                                                    Else
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    End If
                                            End Select
                                    End Select
                            End Select
                        Next
                End Select
            End If

            'ロッド先端オーダーメイド加算価格キー
            If objKtbnStrc.strcSelection.strRodEndOption.Trim <> "" Then
                If InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "WF") = 0 Then
                    decWFLength = 1
                Else
                    For intLoopCnt = InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "WF") + 2 To Len(objKtbnStrc.strcSelection.strRodEndOption.Trim)
                        If Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "0" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "1" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "2" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "3" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "4" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "5" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "6" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "7" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "8" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "9" Or _
                           Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, intLoopCnt, 1) = "." Then
                            If intLoopCnt = Len(objKtbnStrc.strcSelection.strRodEndOption.Trim) Then
                                decLength = intLoopCnt - (InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "WF") + 2) + 1
                            End If
                        Else
                            decLength = intLoopCnt - (InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "WF") + 2) + 1
                            Exit For
                        End If
                    Next

                    decWFLength = CDec(Mid(objKtbnStrc.strcSelection.strRodEndOption.Trim, InStr(1, objKtbnStrc.strcSelection.strRodEndOption.Trim, "WF") + 2, decLength)) - objKtbnStrc.strcSelection.strRodEndWFStdVal
                End If

                Select Case True
                    Case 1 <= decWFLength And decWFLength <= 100
                        strStdWFLength = "100"
                    Case 101 <= decWFLength And decWFLength <= 200
                        strStdWFLength = "200"
                    Case 201 <= decWFLength And decWFLength <= 300
                        strStdWFLength = "300"
                    Case 301 <= decWFLength And decWFLength <= 400
                        strStdWFLength = "400"
                    Case 401 <= decWFLength
                        strStdWFLength = "500"
                End Select
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-TIP-OF-ROD-" & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & strStdWFLength
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            ' オプション外
            If objKtbnStrc.strcSelection.strOtherOption.Trim <> "" Then
                'クッションニードル位置指定の加算
                If Left(objKtbnStrc.strcSelection.strOtherOption.Trim, 1) = "R" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-R-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If

                'ポート2箇所の加算(E)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("E") >= 0 And _
                   objKtbnStrc.strcSelection.strOtherOption.IndexOf("E1") < 0 And _
                   objKtbnStrc.strcSelection.strOtherOption.IndexOf("E2") < 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-E-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If

                'ポート2箇所の加算(E1)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("E1") >= 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-E1-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If

                'ポート2箇所の加算(E2)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("E2") >= 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-E2-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If

                'ポートサイズダウンの加算
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("F") >= 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-F-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    'RM1305007 2013/05/07
                    If objKtbnStrc.strcSelection.strSeriesKataban = "SCS" Or objKtbnStrc.strcSelection.strSeriesKataban = "SCS2" Then
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "", "2", "F"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "B"
                                decOpAmount(UBound(decOpAmount)) = 2
                            Case "D", "G"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Else
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If

                '二山ナックル・二山クレビスの加算(P5)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P5") >= 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-P5-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    'RM1305007 2013/05/07
                    '↓RM1310004 2013/10/01 追加(SCS2)
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "SCS"
                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                Case ""
                                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "CB" And _
                                       objKtbnStrc.strcSelection.strOpSymbol(18).IndexOf("Y") >= 0 Then
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                                Case "B"
                                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Else
                                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "CB" Then
                                            decOpAmount(UBound(decOpAmount)) = 2
                                        Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    End If
                                Case "D"
                                    If objKtbnStrc.strcSelection.strRodEndOption.Trim = "" Then
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                            End Select
                        Case "SCS2"
                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                Case "", "F"
                                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "CB" And _
                                       objKtbnStrc.strcSelection.strOpSymbol(19).IndexOf("Y") >= 0 Then
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                                Case "B"
                                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Else
                                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "CB" Then
                                            decOpAmount(UBound(decOpAmount)) = 2
                                        Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    End If
                                Case "D", "G"
                                    If objKtbnStrc.strcSelection.strRodEndOption.Trim = "" Then
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                            End Select
                        Case Else
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If

                '二山ナックル・二山クレビスの加算(P7)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P7") >= 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-P7-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim

                    'RM1305007 2013/05/07
                    'RM1310004 2013/10/01 追加
                    'If objKtbnStrc.strcSelection.strSeriesKataban = "SCS" Then
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "", "2", "F"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "B"
                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                                decOpAmount(UBound(decOpAmount)) = 2
                            Else
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "D", "G"
                            If objKtbnStrc.strcSelection.strRodEndOption.Trim = "" Then
                                decOpAmount(UBound(decOpAmount)) = 2
                            Else
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
                    'Else
                    '    decOpAmount(UBound(decOpAmount)) = 1
                    'End If
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If

                '二山ナックル・二山クレビスの加算(P8)
                'RM1305007 2013/05/07
                '↓RM1310004 2013/10/01 追加(SCS2)
                If objKtbnStrc.strcSelection.strSeriesKataban = "SCS" Or _
                   objKtbnStrc.strcSelection.strSeriesKataban = "SCS2" Then
                    If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P8") >= 0 Then
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "", "F"
                                If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "CB" Then
                                    'P5
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-P5-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                    End If

                                    'P8
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-P8-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                    End If
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-P8-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                    End If
                                End If
                            Case "B"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-P8-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 2
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                    End If
                                Else
                                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "CB" Then
                                        'P5
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-P5-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If

                                        'P8
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-P8-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-P8-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    End If
                                End If
                            Case "D", "G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-P7-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                If objKtbnStrc.strcSelection.strRodEndOption.Trim = "" Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select
                    End If
                End If

                'タイロッド延長寸法の加算
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("MX") >= 0 Then
                    If InStr(InStr(1, objKtbnStrc.strcSelection.strOtherOption.Trim, "MX") + 1, objKtbnStrc.strcSelection.strOtherOption.Trim, "R") = 0 And _
                       InStr(InStr(1, objKtbnStrc.strcSelection.strOtherOption.Trim, "MX") + 1, objKtbnStrc.strcSelection.strOtherOption.Trim, "H") = 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-MX-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Else
                        ' Rの加算
                        If InStr(InStr(1, objKtbnStrc.strcSelection.strOtherOption.Trim, "MX") + 1, objKtbnStrc.strcSelection.strOtherOption.Trim, "R") <> 0 And _
                           InStr(InStr(1, objKtbnStrc.strcSelection.strOtherOption.Trim, "MX") + 1, objKtbnStrc.strcSelection.strOtherOption.Trim, "R1") = 0 And _
                           InStr(InStr(1, objKtbnStrc.strcSelection.strOtherOption.Trim, "MX") + 1, objKtbnStrc.strcSelection.strOtherOption.Trim, "R2") = 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-MXR-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        End If

                        ' R1の加算
                        If InStr(InStr(1, objKtbnStrc.strcSelection.strOtherOption.Trim, "MX") + 1, objKtbnStrc.strcSelection.strOtherOption.Trim, "R1") <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-MXR1-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        End If

                        ' R2の加算
                        If InStr(InStr(1, objKtbnStrc.strcSelection.strOtherOption.Trim, "MX") + 1, objKtbnStrc.strcSelection.strOtherOption.Trim, "R2") <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-MXR2-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        End If

                        ' Hの加算
                        If InStr(InStr(1, objKtbnStrc.strcSelection.strOtherOption.Trim, "MX") + 1, objKtbnStrc.strcSelection.strOtherOption.Trim, "H") <> 0 And _
                           InStr(InStr(1, objKtbnStrc.strcSelection.strOtherOption.Trim, "MX") + 1, objKtbnStrc.strcSelection.strOtherOption.Trim, "H1") = 0 And _
                           InStr(InStr(1, objKtbnStrc.strcSelection.strOtherOption.Trim, "MX") + 1, objKtbnStrc.strcSelection.strOtherOption.Trim, "H2") = 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-MXH-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        End If

                        ' H1の加算
                        If InStr(InStr(1, objKtbnStrc.strcSelection.strOtherOption.Trim, "MX") + 1, objKtbnStrc.strcSelection.strOtherOption.Trim, "H1") <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-MXH1-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        End If

                        ' H2の加算
                        If InStr(InStr(1, objKtbnStrc.strcSelection.strOtherOption.Trim, "MX") + 1, objKtbnStrc.strcSelection.strOtherOption.Trim, "H2") <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-MXH2-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        End If
                    End If
                End If

                'タイロッド材質ＳＵＳの加算
                'RM1305007 2013/05/07
                If objKtbnStrc.strcSelection.strSeriesKataban = "SCS" Or objKtbnStrc.strcSelection.strSeriesKataban = "SCS2" Then
                    If objKtbnStrc.strcSelection.strOtherOption.IndexOf("M1") >= 0 Then
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "", "F"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-M1-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS1.ToString
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            Case "B"
                                'S1
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-M1-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS1.ToString
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If

                                'S2
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-M1-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS2.ToString
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            Case "D", "G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-M1-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS1.ToString
                                decOpAmount(UBound(decOpAmount)) = 2
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select
                    End If
                End If

                'ピストンロッドはジャバラ付寸法でジャバラ無しの加算
                'RM1305007 2013/05/07
                If objKtbnStrc.strcSelection.strSeriesKataban = "SCS" Or objKtbnStrc.strcSelection.strSeriesKataban = "SCS2" Then
                    If objKtbnStrc.strcSelection.strOtherOption.IndexOf("J9") >= 0 Then
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "", "F"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-J9-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS1.ToString
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            Case "B"
                                'S1
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-J9-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS1.ToString
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If

                                'S2
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-J9-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS2.ToString
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            Case "D", "G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-J9-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           intStrokeS1.ToString
                                decOpAmount(UBound(decOpAmount)) = 2
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select
                    End If
                End If

                'スクレーパ、ロッドパッキンのみフッ素ゴムの加算
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("T9") >= 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-T9-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    'RM1305007 2013/05/07
                    If objKtbnStrc.strcSelection.strSeriesKataban = "SCS" Then
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "", "F"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "B"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "CB" And _
                                       objKtbnStrc.strcSelection.strOpSymbol(18).IndexOf("Y") >= 0 Then
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 2
                                End If
                            Case "D", "G"
                                decOpAmount(UBound(decOpAmount)) = 2
                        End Select
                    ElseIf objKtbnStrc.strcSelection.strSeriesKataban = "SCS2" Then
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "", "F"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "B"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "CB" And _
                                       objKtbnStrc.strcSelection.strOpSymbol(19).IndexOf("Y") >= 0 Then
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 2
                                End If
                            Case "D", "G"
                                decOpAmount(UBound(decOpAmount)) = 2
                        End Select
                    Else
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
