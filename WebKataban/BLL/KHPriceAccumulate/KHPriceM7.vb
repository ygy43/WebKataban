'************************************************************************************
'*  ProgramID  ：KHPriceM8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/07   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：スーパーコンパクトシリンダ　ＳＳＤ－Ｋ
'*
'*  ・受付No：RM0907070  二次電池対応機器　SSD-K
'*                                      更新日：2009/08/20   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'************************************************************************************
Module KHPriceM7

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStrokeS1 As Integer = 0
        Dim intStrokeS2 As Integer = 0
        Dim bolC5Flag As Boolean

        Dim bolOptionM As Boolean = False
        Dim bolOptionN As Boolean = False
        Dim bolOptionP5 As Boolean = False
        Dim bolOptionP51 As Boolean = False
        Dim bolOptionA2 As Boolean = False
        Dim bolOptionP4 As Boolean = False          'RM0907070 2009/08/20 Y.Miura　二次電池対応

        Dim decLength As Decimal
        Dim decWFLength As Decimal
        Dim strStdWFLength As String = String.Empty

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(17), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "M"
                        bolOptionM = True
                    Case "N"
                        bolOptionN = True
                        'RM0907070 2009/08/20 Y.Miura　二次電池対応
                    Case "P4", "P40"
                        bolOptionP4 = True
                    Case "P5"
                        bolOptionP5 = True
                    Case "P51"
                        bolOptionP51 = True
                    Case "A2"
                        bolOptionA2 = True
                End Select
            Next

            'C5チェック
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

            'ストローク設定(S1)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                intStrokeS1 = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                        CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim), _
                                                        CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim))
            End If
            'ストローク設定(S2)
            intStrokeS2 = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                    CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim), _
                                                    CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim))

            '基本価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                'S1
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-K-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                           intStrokeS1.ToString
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If

                'S2
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-K-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                           intStrokeS2.ToString
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            Else
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                    'S1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-K-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS1.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If

                    'S2
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-K-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If
            End If

            'バリエーション加算価格キー
            '(*B*)背合せ形
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-B-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '(*G*)強力スクレーパ形
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G") >= 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G1") < 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G2") < 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G3") < 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G4") < 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G5") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '(*G1*)コイルスクレーパ形
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G1") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G1-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '(*G2*)耐切削油スクレーパ形(一般用)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G2") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    'S1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-KG2-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS1.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If

                    'S2
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-KG2-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Else
                    'S2
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-KG2-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If
            End If

            '(*G3*)耐切削油スクレーパ形(塩素系用)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G3") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    'S1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-KG3-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS1.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If

                    'S2
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-KG3-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-KG3-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If
            End If

            '(*G4*)スパッタ付着防止形
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G4") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G4-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '(*G5*)耐環境スクレーパ形
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G5") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G5-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '(*M*)回り止め形
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("M") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-M-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '(*O*)低速形
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("O") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-O-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '(*T*)耐熱形120℃
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T") >= 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T1") < 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T1L") < 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T2") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-T-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '(*T1*)耐熱形150℃
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T1") >= 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T1L") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-T1-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '(*T1L*)耐熱形スイッチ付
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T1L") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-T1L-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '(*T2*)パッキン材質フッ素ゴム
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T2") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-T2-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '(*U*)低摩擦形
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("U") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-U-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '(*W*)二段形
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-W-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション(M)回り止め加算価格キー
            'S1
            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                Case "M"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-M-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
            End Select

            'S2
            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                Case "M"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-M-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
            End Select

            '微速加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Case "F"
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                        'S1
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "12", "16"
                                Select Case True
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 15
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-15"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 16 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 50
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-16-50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 51
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                            Case "20"
                                Select Case True
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 15
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-15"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 16 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 50
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-16-50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 51 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 100
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 101
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-200"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                            Case "25", "32", "40", "50"
                                Select Case True
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 50
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 51 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 100
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 101 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 150
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-150"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 151 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 200
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-151-200"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 201
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-201-300"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                            Case "63", "80", "100"
                                Select Case True
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 50
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 51 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 100
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 101 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 200
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-200"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 201
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-201-300"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                        End Select

                        'S2
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "12", "16"
                                Select Case True
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 15
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-15"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 16 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 50
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-16-50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 51
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                            Case "20"
                                Select Case True
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 15
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-15"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 16 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 50
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-16-50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 51 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 100
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 101
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-200"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                            Case "25", "32", "40", "50"
                                Select Case True
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 50
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 51 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 100
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 101 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 150
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-150"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 151 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 200
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-151-200"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 201
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-201-300"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                            Case "63", "80", "100"
                                Select Case True
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 50
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 51 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 100
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 101 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 200
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-200"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 201
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-201-300"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                        End Select
                    Else
                        'S2
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "12", "16"
                                Select Case True
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 15
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-15"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 16 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 50
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-16-50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 51
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                            Case "20"
                                Select Case True
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 15
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-15"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 16 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 50
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-16-50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 51 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 100
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 101
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-200"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                            Case "25", "32", "40", "50"
                                Select Case True
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 50
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 51 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 100
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 101 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 150
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-150"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 151 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 200
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-151-200"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 201
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-201-300"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                            Case "63", "80", "100"
                                Select Case True
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 50
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 51 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 100
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 101 And _
                                         CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) <= 200
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-200"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) >= 201
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-KF-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-201-300"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                        End Select
                    End If
            End Select

            'ゴムエアクッション付＆NPTねじ、Gねじ加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                Case "C"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-K-*C-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                        decOpAmount(UBound(decOpAmount)) = 2
                    Else
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Case "GC", "NC"
                    'クッション加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-K-*C-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                        decOpAmount(UBound(decOpAmount)) = 2
                    Else
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
            End Select

            'スイッチ付加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'スイッチ形番＆リード線長さ加算価格キー
            'S1
            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(10).Trim)

                If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                        Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                             "T3V", "T5H", "T5V", "T2YH", "T2YV", _
                             "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                             "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(1)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(10).Trim)
                        Case "T2YD"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(2)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(10).Trim)
                        Case "T2YDT"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(3)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(10).Trim)
                        Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", _
                             "T2YMV", "T3YMH", "T3YMV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(4)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(10).Trim)
                        Case "T2JH", "T2JV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(5)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(10).Trim)
                        Case "ET0H", "ET0V"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(6)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(10).Trim)
                        Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(7)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(10).Trim)
                        Case "V0", "V7"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(8)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(10).Trim)
                    End Select
                End If
            End If

            'S2
            If objKtbnStrc.strcSelection.strOpSymbol(14).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)

                If objKtbnStrc.strcSelection.strOpSymbol(15).Trim <> "" Then
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(14).Trim
                        Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                             "T3V", "T5H", "T5V", "T2YH", "T2YV", _
                             "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                             "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(1)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                        Case "T2YD"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(2)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                        Case "T2YDT"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(3)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                        Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", _
                             "T2YMV", "T3YMH", "T3YMV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(4)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                        Case "T2JH", "T2JV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(5)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                        Case "ET0H", "ET0V"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(6)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                        Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(7)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                        Case "V0", "V7"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(8)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                    End Select
                End If

                'RMXXXXXXX 2009/09/11 Y.Miura　二次電池対応
                If bolOptionP4 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-P4"
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(16).Trim)
                End If
            End If

            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(17), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "M"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "12", "16", "20", "25"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            Case "32", "40", "50", "63", "80", _
                                 "100", "125", "140", "160"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                                    'S1
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                               intStrokeS1.ToString
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                    End If
                                    'S2
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                               intStrokeS2.ToString
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                    End If
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                               intStrokeS2.ToString
                                    decOpAmount(UBound(decOpAmount)) = 1
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                    End If
                                End If
                        End Select

                        '背合せ形＆二段形の時(+α加算する)
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-(B/W)" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        End If
                    Case "N"
                        '￥0
                    Case "S"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                        'RM0907070 2009/08/20 Y.Miura　二次電池対応
                    Case "P4", "P40"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "P5", "P51"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                   Left(strOpArray(intLoopCnt).Trim, 2) & "*" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case "P6"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                   Left(strOpArray(intLoopCnt).Trim, 2) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case "P7", "P71"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                   Left(strOpArray(intLoopCnt).Trim, 2) & "*" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case "P12"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        '2011/1/31 MOD RM1101046(2月VerUP：障害対応(単価キー不正)) START--->
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                        '                                           Left(strOpArray(intLoopCnt).Trim, 2) & CdCst.Sign.Hypen & _
                        '                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        '2011/1/31 MOD RM1101046(2月VerUP：障害対応(単価キー不正)) <---END
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                        '2011/1/13 ADD RM1101046(2月VerUP：SSDシリーズ オプション追加) START--->
                    Case "R1", "R2"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                        '2011/1/13 ADD RM1101046(2月VerUP：SSDシリーズ オプション追加) <---END
                    Case "A2"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                   Left(strOpArray(intLoopCnt).Trim, 2) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                            Select Case True
                                Case objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "N" And bolOptionN = True
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Case objKtbnStrc.strcSelection.strOpSymbol(11).Trim <> "N" And bolOptionN = True
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "N" And bolOptionN = False
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                End Select
            Next

            '支持金具加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(18).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(18).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '付属品加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(19).Trim
                Case "I", "I2", "Y", "Y2"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(19).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                        Select Case True
                            Case objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "N" And bolOptionN = True
                                decOpAmount(UBound(decOpAmount)) = 2
                            Case objKtbnStrc.strcSelection.strOpSymbol(11).Trim <> "N" And bolOptionN = True
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "N" And bolOptionN = False
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Else
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case "IY"
                    'I加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                               Left(objKtbnStrc.strcSelection.strOpSymbol(19).Trim, 1) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    'Y加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                               Right(objKtbnStrc.strcSelection.strOpSymbol(19).Trim, 1) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "I2Y2"
                    'I2加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                               Left(objKtbnStrc.strcSelection.strOpSymbol(19).Trim, 2) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    'Y2加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                               Right(objKtbnStrc.strcSelection.strOpSymbol(19).Trim, 2) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'ロッド先端オーダーメイド加算価格キー
            'If objKtbnStrc.strcSelection.strOpSymbol(20).Trim <> "" Then
            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            '    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-TIP-OF-ROD-" & _
            '                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
            '    decOpAmount(UBound(decOpAmount)) = 1
            '    If bolC5Flag = True Then
            '        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
            '    End If
            'End If

            'ロッド先端特注加算価格キー
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
                    Case 0 <= decWFLength And decWFLength <= 100
                        strStdWFLength = "100"
                    Case 101 <= decWFLength And decWFLength <= 200
                        strStdWFLength = "200"
                    Case 201 <= decWFLength And decWFLength <= 300
                        strStdWFLength = "300"
                    Case 301 <= decWFLength And decWFLength <= 400
                        strStdWFLength = "400"
                    Case 401 <= decWFLength And decWFLength <= 500
                        strStdWFLength = "500"
                    Case 501 <= decWFLength And decWFLength <= 600
                        strStdWFLength = "600"
                    Case 601 <= decWFLength
                        strStdWFLength = "700"
                End Select
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-TIP-OF-ROD-" & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & strStdWFLength
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If

            End If


        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
