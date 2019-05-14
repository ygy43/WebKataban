'************************************************************************************
'*  ProgramID  ：KHPrice79
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/22   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：スーパーツインロッドシリンダ
'*             ：ＳＴＲ２－Ｂ　ベアリング軸受形
'*             ：ＳＴＲ２－Ｍ　メタル軸受形
'*
'*  ・受付No：RM0908030  二次電池対応機器　STR2
'*                                      更新日：2009/10/15   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'************************************************************************************
Module KHPrice79

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer
        Dim strOptionP4 As String = String.Empty            'RM0908030 2009/10/15 Y.Miura　二次電池対応
        Dim bolC5Flag As Boolean                            'RM0908030 2009/10/15 Y.Miura　二次電池対応

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)            'RM0908030 2009/10/30 Y.Miura

            'ストロークを設定する
            intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim), _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim))

            '二次電池判定                    'RM0908030 2009/10/15 Y.Miura
            Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                Case "P4", "P40"
                    strOptionP4 = objKtbnStrc.strcSelection.strOpSymbol(9).Trim
            End Select

            'C5チェック                      'RM0908030 2009/10/15 Y.Miura
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc)

            'バリエーション加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "O", "Q"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    'RM0908030 2009/10/30 Y.Miura
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Case "F"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 50 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & "STR10" & CdCst.Sign.Hypen & "50"
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0908030 2009/10/30 Y.Miura
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & "STR51" & CdCst.Sign.Hypen & "100"
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0908030 2009/10/30 Y.Miura
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    End If
            End Select

            '基本価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "D"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    'RM0908030 2009/10/30 Y.Miura
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    'RM0908030 2009/10/30 Y.Miura
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
            End Select

            'スイッチ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)

                'リード線長さ加算価格キー
                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                End If

                'P4加算
                'RM0908030 2009/10/15 Y.Miura
                If strOptionP4 <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SW-P4"
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                End If
            End If

            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                        'RM0908030 2009/10/15 Y.Miura
                        'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                        '                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                        '                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                   "-OP-" & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0908030 2009/10/30 Y.Miura
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                End Select
            Next

            'クリーン仕様加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                'RM0908030 2009/10/15 Y.Miura
                'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                '                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim & CdCst.Sign.Hypen & _
                '                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-OP-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                'RM0908030 2009/10/30 Y.Miura
                'RM0912039 2009/12/18 Y.Miura クリーン仕様はC5加算なし
                'If bolC5Flag = True Then
                '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                'End If
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
