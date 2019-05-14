'************************************************************************************
'*  ProgramID  ：KHPriceB0
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/23   作成者：NII K.Sudoh
'*
'*  概要       ：タイトシリンダ　ＣＭＫ２　基本ベース
'*
'*  更新履歴   ：                       更新日：2007/05/16   更新者：NII A.Takahashi
'*               ・T2W/T3Wスイッチ追加に伴い、リード線加算ロジック部を修正
'*                                      更新日：2008/03/24   更新者：T.Sato
'*               ・受付No.RM0707057 CMK2ロッド先端特注対応
'*  　             WFサイズから単価キーを作成するロジックを追加
'*  ・受付No：RM0908030  二次電池対応機器　
'*                                      更新日：2009/09/04   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'************************************************************************************
Module KHPriceB0

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStrokeS1 As Integer = 0
        Dim intStrokeS2 As Integer = 0
        Dim bolC5Flag As Boolean
        Dim strOptionP4 As String = String.Empty        'RM0908030 2009/09/04 Y.Miura　二次電池対応

        Dim decLength As Decimal
        Dim decWFLength As Decimal
        Dim strStdWFLength As String = String.Empty

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'オプション加算価格キー
            'RM0908030 2009/09/04 Y.Miura　二次電池対応
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(15), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "P4", "P40"
                        strOptionP4 = strOpArray(intLoopCnt).Trim
                End Select
            Next

            'C5チェック
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

            'ストローク取得(S1)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                intStrokeS1 = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                      CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                      CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim))
            End If
            'ストローク取得(S2)
            intStrokeS2 = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim))

            '基本価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                'S1
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

                'S2
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

            Else
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Then
                    'ストローク調整形の場合はDロッドベースにて価格算出
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-BASE-D-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Else
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
                End If

                'RM0908030 2009/09/04 Y.Miura　二次電池対応
                If strOptionP4 <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                               strOptionP4 & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

            End If

            'バリエーション加算価格キー
            '背合せ形(*B*)
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
            'エアクッション形(*C*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("C") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-C-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If
            '微速形(*F*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("F") >= 0 Then
                'S1
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    Select Case True
                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 50
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-F-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-15-50"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) >= 51 And _
                             CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 300
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-F-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-51-300"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) >= 301
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-F-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-301-750"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                    End Select
                End If

                'S2
                Select Case True
                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) <= 50
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-F-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-15-50"
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) >= 51 And _
                         CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) <= 300
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-F-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-51-300"
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) >= 301
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-F-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-301-750"
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                End Select
            End If
            '強力スクレーパ形(*G*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G") >= 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G1") < 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G2") < 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G3") < 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G4") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-G-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If
            'コイルスクレーパ形(*G1*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G1") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-G1-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If
            '耐切削油スクレーパ形(*G2*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G2") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "JG2" Then
                    'S1
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-JG2-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                   intStrokeS1.ToString
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'S2
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-JG2-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    'S1
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-G2-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                   intStrokeS1.ToString
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    End If

                    'S2
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-G2-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If
               
            End If
            '耐切削油スクレーパ形(*G3*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G3") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "JG3" Then
                    'S1
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-JG3-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                   intStrokeS1.ToString
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'S2
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-JG3-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    'S1
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-G3-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                   intStrokeS1.ToString
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    End If

                    'S2
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-G3-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                               intStrokeS2.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If
               
            End If
            'スパッタ付着防止形(*G4*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G4") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-G4-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If
            '低油圧形(*H*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("H") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-H-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If
            '回り止め形(*M*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("M") >= 0 Then
                Select Case True
                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) <= 300
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-M-" & _
                                                            objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-25-300"
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) >= 301
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-M-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-301-500"
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                End Select
            End If
            '低速形(*O*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("O") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-O-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If
            'ストローク調整形(押出し)(*P*)
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
            '落下防止形(*Q*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-Q-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If
            'ストローク調整形(引込み)(*R*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("R") >= 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("SR") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-R-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If
            '単動押出し形(*S*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("S") >= 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("SR") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-S-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If
            '単動引込み形(*SR*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("SR") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-SR-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If
            '耐熱形(*T*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T") >= 0 And _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T2") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-T-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If
            'パッキン材質フッ素ゴム(*T2*")
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T2") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-T2-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If
            'スピードコントローラ内臓形(*Z*)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Z") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-VAR-Z-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '支持形式価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "JG2" Or objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "JG3" Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    Case "LB", "LS", "FA", "FB", "CC"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SUPPORT-JG-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'CC/CC1の場合は率計算
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            Case "CC", "CC1"
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select
                End Select
            Else
                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    Case "00", "LB", "LS", "FA", "FB"
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SUPPORT-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        'CC/CC1の場合は率計算
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            Case "CC", "CC1"
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select
                End Select
            End If

            'ゴムエアクッション加算価格キー
            If Right(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "C" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-*C-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                    decOpAmount(UBound(decOpAmount)) = 2
                Else
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'スイッチ形番加算価格キー
            'S1
            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SW-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)

                'リード線長さ加算価格キー
                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    Case "3", "5"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            Case "T2H", "T2V", "T2YH", "T2YV", "T3H", _
                                 "T3V", "T3YH", "T3YV", "T0H", "T0V", _
                                 "T5H", "T5V", "T1H", "T1V", "T8H", "T8V", _
                                 "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(1)-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                            Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", _
                                 "T2YMV", "T3YMH", "T3YMV"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(2)-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                            Case "T2JH", "T2JV"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(3)-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                            Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(4)-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                        End Select
                End Select
            End If
            'S2
            If objKtbnStrc.strcSelection.strOpSymbol(12).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SW-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(14).Trim)

                'RM0908030 2009/09/04 Y.Miura 二次電池対応
                If strOptionP4 <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SW-P4"
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(14).Trim)
                End If

                'リード線長さ加算価格キー
                Select Case objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                    Case "3", "5"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                            Case "T2H", "T2V", "T2YH", "T2YV", "T3H", _
                                 "T3V", "T3YH", "T3YV", "T0H", "T0V", _
                                 "T5H", "T5V", "T1H", "T1V", "T8H", "T8V", _
                                 "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(1)-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(14).Trim)
                            Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", _
                                 "T2YMV", "T3YMH", "T3YMV"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(2)-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(14).Trim)
                            Case "T2JH", "T2JV"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(3)-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(14).Trim)
                            Case "T2YLH", "T2YLV", "T3YLH", "T3YLV"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SWLW(4)-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(14).Trim)
                        End Select
                End Select
            End If

            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(15), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "J", "L", "M"
                        'S1
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
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
                        'RM0908030 2009/09/04 Y.Miura 二次電池対応
                    Case "P4", "P40"

                    Case "P6", "A2"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
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
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-P7*-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case "M0", "M1"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case "F", "FE"
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("S") >= 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & "S"
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                       strOpArray(intLoopCnt).Trim
                        End If
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                End Select
            Next

            '付属品加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(16), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "IY"
                        'I加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-I"
                        decOpAmount(UBound(decOpAmount)) = 1

                        'Y加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-Y"
                        decOpAmount(UBound(decOpAmount)) = 1

                    Case "FP1"
                        '食品製造工程向け商品
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            Case "B"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                  strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "B"
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                                  strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select

                    Case Else
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            Case "JG2", "JG3"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-JG-" & _
                                                                     strOpArray(intLoopCnt).Trim
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-" & _
                                                                      strOpArray(intLoopCnt).Trim
                        End Select

                        Select Case strOpArray(intLoopCnt).Trim
                            Case "I", "Y"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            Case "B2"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select
            Next

            '付属品加算価格キー
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "5"
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(17), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "JG2", "JG3"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-JG-" & _
                                                                                   strOpArray(intLoopCnt).Trim
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-ACC-" & _
                                                                                   strOpArray(intLoopCnt).Trim
                                End Select

                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "I", "Y"
                                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") >= 0 Then
                                            decOpAmount(UBound(decOpAmount)) = 2
                                        Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    Case "B2"
                                        decOpAmount(UBound(decOpAmount)) = 1

                                End Select
                        End Select
                    Next
                Case Else
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
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-TIP-OF-ROD-" & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & strStdWFLength
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
