'************************************************************************************
'*  ProgramID  ：KHPrice72
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/22   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：スーパーロッドレスシリンダ　ＳＲＭ／ＳＲＭ－Ｑ
'*
'*  更新       ：　　
'*   RM0811134 2009/05/21 Y.Miura SRM3シリーズの追加
'*                                      更新日：2009/07/22   更新者：Y.Miura
'*   RM0906066 2009/07/23 Y.Miura SRM3 C0,C1オプション追加
'*                                      更新日：2009/07/23   更新者：Y.Miura
'*  ・受付No：RM0908030  二次電池対応機器　SRM3
'*                                      更新日：2009/10/15   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'************************************************************************************
Module KHPrice72

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
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(1).Trim), _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim))
            'RM0811134 2009/05/21 Y.Miura
            'シリーズ形番の第1ハイフン前を取得
            Dim strHySeriesKataban As String = objKtbnStrc.strcSelection.strSeriesKataban.Trim
            If InStr(strHySeriesKataban, "-") > 0 Then
                strHySeriesKataban = strHySeriesKataban.Substring(0, InStr(strHySeriesKataban, "-") - 1)
            End If

            'オプションの要素番号
            Dim intOption As Integer
            If strHySeriesKataban = "SRM3" Then
                intOption = 8
            Else
                intOption = 7
            End If

            '二次電池判定(SRM3のみ)                    'RM0908030 2009/10/15 Y.Miura
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intOption), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "P4", "P40"
                        strOptionP4 = strOpArray(intLoopCnt).Trim
                End Select
            Next

            'C5チェック                      'RM0908030 2009/10/15 Y.Miura
            'RM1001043 2010/02/22 Y.Miura 二次電池C5加算廃止
            'bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc)
            'bolC5Flag = False

            'RM1306001 2013/06/06
            'C5チェック
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
            'RM0811134 2009/05/21 Y.Miura
            'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
            '                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
            '                                           intStroke.ToString
            strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                       intStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1
            'RM0908030 2009/10/30 Y.Miura
            If bolC5Flag = True Then
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
            End If


            '落下防止機構加算価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                'Case "SRM-Q"
                Case "SRM-Q", "SRM3-Q"          'Y.Miura 2009/05/21
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    'RM0908030 2009/10/30 Y.Miura
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
            End Select

            'スイッチ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                'RM0811134 2009/05/21 Y.Miura
                'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                '                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                'strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & CdCst.Sign.Hypen & _
                '                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                'RM0908030 2009/10/19 Y.Miura
                strOpRefKataban(UBound(strOpRefKataban)) = Left(strHySeriesKataban, 3) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)

                'P4加算
                'RM0908030 2009/10/15 Y.Miura
                If strOptionP4 <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(strHySeriesKataban, 3) & "-SW-P4"
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                End If

                'リード線長さ加算価格キー
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                    Select Case Mid(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 4, 1)
                        Case "F", "M"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                            'RM0811134 2009/05/21 Y.Miura
                            'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                            '                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "FM"
                            'strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & CdCst.Sign.Hypen & _
                            '                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "FM"
                            'RM0908030 2009/10/19 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(strHySeriesKataban, 3) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "FM"
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                        Case "D"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                            'RM0811134 2009/05/21 Y.Miura
                            'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                            '                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                            '                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            'strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & CdCst.Sign.Hypen & _
                            '                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                            '                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            'RM0908030 2009/10/19 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(strHySeriesKataban, 3) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                            'RM0811134 2009/05/21 Y.Miura
                            'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                            '                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            'strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & CdCst.Sign.Hypen & _
                            '                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            'RM0908030 2009/10/19 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(strHySeriesKataban, 3) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                    End Select
                End If
            End If

            'RM0906066 2009/07/23 Y.Miura 追加 ↓↓
            If strHySeriesKataban = "SRM3" Then
                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                    strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    'RM0908030 2009/10/30 Y.Miura
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If
            End If
            'RM0906066 2009/07/23 Y.Miura 追加 ↑↑

            'オプション加算価格キー
            'RM0906066 2009/07/23 Y.Miura 変更 
            'strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intOption), CdCst.Sign.Delimiter.Comma)

            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                        'RM0811134 2009/05/21 Y.Miura
                        'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                        '                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                        '                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        'strOpRefKataban(UBound(strOpRefKataban)) = strHySeriesKataban & CdCst.Sign.Hypen & _
                        '                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                        '                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        'RM0908030 2009/10/19 Y.Miura
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strHySeriesKataban, 3) & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim

                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0908030 2009/10/30 Y.Miura
                        'RM0912XXX 2009/12/09 Y.Miura　二次電池以外はC5加算
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "P4", "P40"
                            Case Else
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                        End Select

                        'RM0908030 2009/10/15 Y.Miura　二次電池対応
                        If strOptionP4 <> "" Then
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "A", "A1", "A2"    'ショックキラー付
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(strHySeriesKataban, 3) & "-A-" & _
                                                                               strOptionP4 & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case "A"        '２ヶ付
                                            decOpAmount(UBound(decOpAmount)) = 2
                                        Case Else       '１ヶ付
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Case "A3"               'ショックキラー無し
                                Case "E", "E1", "E2"    '軽荷重ショックキラー付
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/30 Y.Miura
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(strHySeriesKataban, 3) & "-E-" & _
                                                                               strOptionP4 & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case "E"        '２ヶ付
                                            decOpAmount(UBound(decOpAmount)) = 2
                                        Case Else       '１ヶ付
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                            End Select
                        End If

                End Select
            Next

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
