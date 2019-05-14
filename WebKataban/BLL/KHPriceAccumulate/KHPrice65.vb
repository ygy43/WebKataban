'************************************************************************************
'*  ProgramID  ：KHPrice65
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/27   作成者：NII K.Sudoh
'*
'*  概要       ：タイロッドシリンダ　ＳＣＧ
'*
'*  更新履歴   ：                       更新日：2007/05/16   更新者：NII A.Takahashi
'*               ・T2W/T3Wスイッチ追加に伴い、リード線加算ロジック部を修正
'*  ・受付No：RM0907070  二次電池対応機器　SCG
'*                                      更新日：2009/08/21   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'************************************************************************************
Module KHPrice65

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)



        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer = 0
        Dim bolC5Flag As Boolean
        Dim bolOptionP4 As Boolean = False          'RM0907070 2009/08/21 Y.Miura　二次電池対応

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'C5チェック
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

            'RM1805007_SCWR,S機種追加対応
            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWR" Or objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWS" Then
                'ストローク設定
                intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                      CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                      CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim))
            Else
                'ストローク設定
                intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                      CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim), _
                                                      CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim))
            End If

            '基本価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCG-D" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-D-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           intStroke.ToString
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
                'RM1712042_SCWP2,SCWT2追加
            ElseIf objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWP2" Or objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWT2" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "*2-" & objKtbnStrc.strcSelection.strKeyKataban & "-BASE-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           intStroke.ToString
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
                'RM1805007_SCWR,S機種追加対応
            ElseIf objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWR" Or objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWS" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                           intStroke.ToString
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           intStroke.ToString
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'RM1805007_SCWR,S機種追加対応
            'バリエーション「L」加算価格キー
            If (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWR" Or _
                objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWS") And _
                objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "L" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「Q」加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("-Q") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-Q-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「G」加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("-G") >= 0 And _
               objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("-G2") < 0 And _
               objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("-G3") < 0 And _
               objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("-G4") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「G4」加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("-G4") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G4-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「G2」加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("-G2") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G2-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           intStroke.ToString
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「G3」加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("-G3") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G3-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           intStroke.ToString
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「O」加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("-O") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-O-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「U」加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("-U") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-U-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「M」加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("-M") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-M-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           intStroke.ToString
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '支持形式加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "00"
                Case "TA", "TB", "TC"
                    'RM1712042_SCWP2,SCWT2追加
                    If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWP2" Or objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWT2" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "*2-" & objKtbnStrc.strcSelection.strKeyKataban & "-SUPPORT-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SUPPORT-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Case Else
                    'RM1712042_SCWP2,SCWT2追加
                    If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWP2" Or objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWT2" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "*2-" & objKtbnStrc.strcSelection.strKeyKataban & "-SUPPORT-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM1805007_SCWR,S機種追加対応
                    ElseIf objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWR" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "00" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    ElseIf objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWS" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SUPPORT-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "J", "M"
                        'RM1712042_SCWP2,SCWT2追加
                        If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWP2" Or objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWT2" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "*2-" & objKtbnStrc.strcSelection.strKeyKataban & "-OP-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                       intStroke.ToString
                            'RM1805007_SCWR,S機種追加対応
                            'RM1807055_口径、ストローク追加
                        ElseIf objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWS" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim

                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                       intStroke.ToString

                        End If
                        If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCG-D" Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                        'RM0907070 2009/08/21 Y.Miura　二次電池対応
                    Case "P4", "P40"
                        bolOptionP4 = True
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCG-D" Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case "P6"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCG-D" Then
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
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0912XXX 2009/12/09 Y.Miura　C5加算もれ
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                End Select
            Next

            '付属品加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(11), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "I", "Y"
                        'RM1712042_SCWP2,SCWT2追加
                        If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWP2" Or objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWT2" Then

                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "*2-" & objKtbnStrc.strcSelection.strKeyKataban & "-ACC-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim

                        Else

                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim

                        End If
                        If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCG-D" Then
                            decOpAmount(UBound(decOpAmount)) = 2
                        Else
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case "IY"
                        'I加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-I-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        'Y加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-Y-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "FP1"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" _
                                                                 & objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If

                    Case Else
                        'RM1712042_SCWP2,SCWT2追加
                        If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWP2" Or objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWT2" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "*2-" & objKtbnStrc.strcSelection.strKeyKataban & "-ACC-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                            'RM1805007_SCWR,S機種追加対応
                        ElseIf objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWS" Then
                            '無し（支持形式加算）
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                       strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                End Select
            Next

            '食品製造工程向け商品
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "F"
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(12), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "I", "Y"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCG-D" Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            Case "IY"
                                'I加算
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-I-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1

                                'Y加算
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-Y-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
            End Select

            'スイッチ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then

                'RM1712042_SCWP2,SCWT2追加
                If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWP2" Or objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWT2" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strKeyKataban & "-SW-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)

                    'リード線長さ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            Case "T1H", "T1V", "T2H", "T2V", "T2YH", "T2YV", "T2WH", "T2WV", "T3H", "T3V", _
                                 "T3YH", "T3YV", "T3WH", "T3WV", "T0H", "T0V", "T5H", "T5V", "T8H", "T8V", _
                                 "T2JH", "T2JV", "T3PH", "T3PV"

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "*2-" & objKtbnStrc.strcSelection.strKeyKataban & "-SWLW(1)-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                'Case "T2YD"
                                '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                '    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(2)-" & _
                                '                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                '    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                            Case "T2YD"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "*2-" & objKtbnStrc.strcSelection.strKeyKataban & "-SWLW(3)-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                            Case "T2YDT"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "*2-" & objKtbnStrc.strcSelection.strKeyKataban & "-SWLW(4)-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        End Select
                    End If
                    'RM1805007_SCWR,S機種追加
                ElseIf objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWR" Or objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWS" Then

                    'RM1807055_SCWR口径追加
                    If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SCWR" Then

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim    '口径
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)

                    Else

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)

                    End If

                    'リード線長さ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            Case "T1H", "T1V", "T2H", "T2V", "T2YH", "T2YV", "T2WH", "T2WV", "T3H", "T3V", _
                                 "T3YH", "T3YV", "T3WH", "T3WV", "T0H", "T0V", "T5H", "T5V", "T8H", "T8V"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-LEAD01-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                            Case "T2YD"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-LEAD02-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                            Case "T2YDT"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-LEAD03-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                            Case "T2JH", "T2JV", "T3PH", "T3PV"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-LEAD04-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        End Select
                    End If

                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)

                    '↓2013/06/27 グローバル機種対応(SCW)
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                        Case "SCW"
                            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim.Length <> 0 Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                    Case "T0H", "T0V"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(1)-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                    Case "T5H", "T5V"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(2)-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                    Case "T2H", "T2V"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(3)-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                    Case "T3H", "T3V"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(4)-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                End Select
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                            End If
                        Case Else
                            'リード線長さ加算価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                    Case "T2H", "T2V", "T2YH", "T2YV", "T3H", _
                                         "T3V", "T3YH", "T3YV", "T0H", "T0V", _
                                         "T5H", "T5V", "T2YLH", "T3YLH", "T2YLV", "T3YLV", _
                                         "T1H", "T1V", "T8H", "T8V", "T2WH", "T2WV", "T3WH", "T3WV", _
                                         "T3PH", "T3PV"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(1)-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                    Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", _
                                         "T2YMV", "T3YMH", "T3YMV"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(2)-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                    Case "T2YD"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(3)-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                    Case "T2YDT"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(4)-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                    Case "T2JH", "T2JV"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(5)-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                End Select
                            End If
                    End Select

                End If

                'RM0907070 2009/08/21 Y.Miura　二次電池対応
                If bolOptionP4 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-P4"
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                End If
            End If

            'スズキ向け特注
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "SCG-G1", "SCG-G1L2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "S"
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCG-TS-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                            End If
                    End Select
            End Select
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "SCG-G1", "SCG-G1L2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "R", "S"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "SCG" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(13).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
