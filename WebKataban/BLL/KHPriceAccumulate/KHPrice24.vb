'************************************************************************************
'*  ProgramID  ：KHPrice24
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/21   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：小形ダイレクトシリンダ　ＭＤＣ２
'*
'*  ・受付No：RM0908030  二次電池対応機器　
'*                                      更新日：2009/09/04   更新者：Y.Miura
'*  ・受付No：RM0912XXX  二次電池対応機器のC5価格適用
'*                                      更新日：2009/12/09   更新者：Y.Miura
'*  ・受付No：RM1002XXX  二次電池対応機器のチェック区分変更
'*                                      更新日：2009/12/09   更新者：Y.Miura
'************************************************************************************
Module KHPrice24

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer

        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False
        Dim bolOptionP4 As Boolean = False      'RM0908030 2009/09/04 Y.Miura　二次電池対応
        Dim bolC5Flag As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)                        'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応

            'RM0908030 2009/09/04 Y.Miura　二次電池対応
            'RM0912XXX 2009/12/08 Y.Miura 不具合対応 要素数が7未満の機種は二時電池ではない
            If objKtbnStrc.strcSelection.strOpSymbol.Length >= 8 Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    Case "P4", "P40"
                        bolOptionP4 = True
                End Select
            End If

            'C5チェック
            'RM1002XXX 2010/02/12 Y.Miura C5加算しない、チェック区分変更
            'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
            'RM1404033 2014/04/11 シリンダスイッチC5加算
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)
            'bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc)

            'ストローク取得
            intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim), _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim))

            'バリエーション加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "F" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            '基本価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "MDC2"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Case "MDC2-L"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Case "MDC2-X"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Case "MDC2-XL"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 6) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Case "MDC2-Y"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Case "MDC2-YL"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 6) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
            End Select

            'マグネット内蔵(L)加算価格キー
            If Right(objKtbnStrc.strcSelection.strSeriesKataban, 1) = "L" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & "L"
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'スイッチ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)

                'RM0908030 2009/09/04 Y.Miura　二次電池対応
                If bolOptionP4 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & "-SW-P4"
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                End If

                'リード線長さ加算価格価格キー
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                End If
            End If

            'クリーン仕様加算価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "MDC2", "MDC2-L"
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        'RM0908030 2009/09/04 Y.Miura　二次電池対応
                        'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & CdCst.Sign.Hypen & _
                        '                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                        '                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & "-OP-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
