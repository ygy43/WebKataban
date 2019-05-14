'************************************************************************************
'*  ProgramID  ：KHPriceM8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/07   作成者：NII K.Sudoh
'*
'*  概要       ：スーパーコンパクトシリンダ　ＳＳＤ
'*
'*  更新履歴   ：                       更新日：2007/05/16   更新者：NII A.Takahashi
'*               ・T2W/T3Wスイッチ追加に伴い、リード線加算ロジック部を修正
'*                                      更新日：2007/10/23   更新者：NII A.Takahashi
'*               ・ロッド先端特注画面の追加により、ロッド先端特注価格ロジック変更
'*  ・受付No：RM0906034  二次電池対応機器　SSD
'*                                      更新日：2009/08/05   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'************************************************************************************
Module KHPriceM6

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStrokeS1 As Integer = 0
        Dim intStrokeS2 As Integer = 0
        Dim bolC5Flag As Boolean

        Dim bolOptionN As Boolean = False
        Dim bolOptionP5 As Boolean = False
        Dim bolOptionP51 As Boolean = False
        Dim bolOptionA2 As Boolean = False
        Dim bolOptionP4 As Boolean = False              'RM0906034 2009/08/05 Y.Miura　二次電池対応 追加

        Dim decLength As Decimal
        Dim decWFLength As Decimal
        Dim strStdWFLength As String = String.Empty

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(19), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "N"
                        bolOptionN = True
                    Case "P4", "P40"                    'RM0906034 2009/08/05 Y.Miura　二次電池対応 追加
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
                                                    CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim))

            '基本価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                'S1
                If objKtbnStrc.strcSelection.strOpSymbol(6).IndexOf("K") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q") >= 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-BASE-K-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS1.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-BASE-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS1.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If

                'S2
                If objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("K") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q") >= 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-BASE-K-" & _
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
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-BASE-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If
            Else
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q") >= 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-BASE-K-" & _
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
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-BASE-" & _
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
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-B-" & _
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
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-G-" & _
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
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-G1-" & _
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
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-G2-" & _
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
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-G2-" & _
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
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-G2-" & _
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
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-G3-" & _
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
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-G3-" & _
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
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-G3-" & _
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
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-G4-" & _
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

            '(*G5*)スパッタ付着防止形
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G5") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-G5-" & _
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
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-M-" & _
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
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-O-" & _
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

            '(*Q*)落下防止形
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-Q-" & _
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

            '(*T*)耐熱形120℃
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T") >= 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T1") < 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T1L") < 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T2") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-T-" & _
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
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-T1-" & _
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
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-T1L-" & _
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

            '(*T2*)パッキン材質フッ素ゴム
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T2") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-T2-" & _
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

            '(*W*)二段形
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-W-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '(*X*)押出し形
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("X") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-X-" & _
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

            '(*Y*)引込み形
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Y") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-Y-" & _
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

            'バリエーション(M)回り止め加算価格キー
            'S1
            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                Case "M", "KM"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-M-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
            End Select
            'S2
            Select Case objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                Case "M", "KM"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-VAR-M-" & _
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
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("K") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(6).IndexOf("K") >= 0 Then
                            '内径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "12", "16"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
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
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-16-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 51
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
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
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
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
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
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
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 101
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
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
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
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
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
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
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
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
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-151-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
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
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
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
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
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
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                            End Select
                        Else
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "12", "16", "20"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 16
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-16-30"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                                Case "25", "32", "40", "50", "63", "80", "100"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 25
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-25"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 26
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-26-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                                Case "125", "140", "160"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
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
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
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
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                            End Select
                        End If

                        'S2
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("K") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("K") >= 0 Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "12", "16"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 16 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-16-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 51
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                                Case "20"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 16 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-16-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 51 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 101
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                                Case "25", "32", "40", "50"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 51 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 101 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 150
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-150"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 151 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-151-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                                Case "63", "80", "100"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 51 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 101 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                            End Select
                        Else
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "12", "16", "20"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 16
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-16-30"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                                Case "25", "32", "40", "50", "63", "80", "100"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 25
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-25"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 26
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-26-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                                Case "125", "140", "160"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 51 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 101 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                            End Select
                        End If
                    Else
                        'S2
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("K") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("K") >= 0 Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "12", "16"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 16 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-16-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 51
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                                Case "20"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 16 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-16-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 51 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 101
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                                Case "25", "32", "40", "50"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 51 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 101 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 150
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-150"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 151 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-151-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                                Case "63", "80", "100"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 51 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 101 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-KF-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                            End Select
                        Else
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "12", "16", "20"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 16
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-16-30"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                                Case "25", "32", "40", "50", "63", "80", "100"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 25
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-25"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 26
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-26-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                                Case "125", "140", "160"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-5-50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 51 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-51-100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 101 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) <= 200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-101-200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) >= 201
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-F-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-201-300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                            End Select
                        End If
                    End If
            End Select

            'NPTねじ、Gねじ加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                Case "GD", "ND"
                    'D加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SCREW-" & _
                                                               Right(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 1) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Case "D"
                    'D加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SCREW-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
            End Select

            'スイッチ付加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SW-" & _
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
            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SW-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(11).Trim)

                If objKtbnStrc.strcSelection.strOpSymbol(10).Trim <> "" Then
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                        Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                             "T3V", "T5H", "T5V", "T2YH", "T2YV", _
                             "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                             "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(1)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(11).Trim)
                        Case "T2YD"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(2)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(11).Trim)
                        Case "T2YDT"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(3)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(11).Trim)
                        Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", _
                             "T2YMV", "T3YMH", "T3YMV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(4)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(11).Trim)
                        Case "T2JH", "T2JV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(5)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(11).Trim)
                        Case "ET0H", "ET0V"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(6)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(11).Trim)
                        Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(7)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(11).Trim)
                        Case "V0", "V7"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(8)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(11).Trim)
                    End Select
                End If
            End If

            'S2
            If objKtbnStrc.strcSelection.strOpSymbol(16).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SW-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(18).Trim)

                If objKtbnStrc.strcSelection.strOpSymbol(17).Trim <> "" Then
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                        Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                             "T3V", "T5H", "T5V", "T2YH", "T2YV", _
                             "T3YH", "T3YV", "T1H", "T1V", "T8H", "T8V", _
                             "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(1)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(18).Trim)
                        Case "T2YD"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(2)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(18).Trim)
                        Case "T2YDT"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(3)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(18).Trim)
                        Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", _
                             "T2YMV", "T3YMH", "T3YMV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(4)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(18).Trim)
                        Case "T2JH", "T2JV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(5)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(18).Trim)
                        Case "ET0H", "ET0V"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(6)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(18).Trim)
                        Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(7)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(18).Trim)
                        Case "V0", "V7"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SWLW(8)-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(17).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(18).Trim)
                    End Select
                End If
                'RM0906034 2009/08/05 Y.Miura　二次電池対応
                If bolOptionP4 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SW-P4"
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(18).Trim)
                End If
            End If

            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(19), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "M"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "12", "16", "20", "25"
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
                            Case "32", "40", "50", "63", "80", _
                                 "100", "125", "140", "160"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                                    'S1
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
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
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
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
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
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
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-(B/W)" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        End If
                    Case "M1"
                        '背合せ形＆二段形の時
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                            'S1
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
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
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
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
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                       intStrokeS2.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        End If

                        '背合せ形＆二段形の時(+α加算する)
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP(B/W)" & _
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
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                        'RM0906034 2009/08/05 Y.Miura　二次電池対応
                    Case "P4", "P40"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "P5", "P51"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
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
                        '2011/1/13 MOD RM1101046(2月VerUP：SSDシリーズ オプション追加) START--->
                    Case "P6", "R1", "R2"
                        'Case "P6"
                        '2011/1/13 MOD RM1101046(2月VerUP：SSDシリーズ オプション追加) <---END
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
                    Case "P7", "P71"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
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
                    Case "A2"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
                                                                   Left(strOpArray(intLoopCnt).Trim, 2) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                            Select Case True
                                Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim = "N" And bolOptionN = True
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim <> "N" And bolOptionN = True
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim = "N" And bolOptionN = False
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
            If objKtbnStrc.strcSelection.strOpSymbol(20).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(20).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '付属品加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(21).Trim
                Case "I", "I2", "Y", "Y2"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-ACC-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(21).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                        Select Case True
                            Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim = "N" And bolOptionN = True
                                decOpAmount(UBound(decOpAmount)) = 2
                            Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim <> "N" And bolOptionN = True
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim = "N" And bolOptionN = False
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
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-ACC-" & _
                                                               Left(objKtbnStrc.strcSelection.strOpSymbol(21).Trim, 1) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    'Y加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-ACC-" & _
                                                               Right(objKtbnStrc.strcSelection.strOpSymbol(21).Trim, 1) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "I2Y2"
                    'I2加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-ACC-" & _
                                                               Left(objKtbnStrc.strcSelection.strOpSymbol(21).Trim, 2) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    'Y2加算
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-ACC-" & _
                                                               Right(objKtbnStrc.strcSelection.strOpSymbol(21).Trim, 2) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

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
