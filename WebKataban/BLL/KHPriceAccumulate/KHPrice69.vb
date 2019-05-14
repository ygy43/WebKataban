'************************************************************************************
'*  ProgramID  ：KHPrice69
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/20   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：小形クロスローラ平行ハンド(ＢＨＡ)
'*             ：センタリングハンド(ＢＨＥ)
'*             ：小形クロスローラ平行ハンド(ＢＨＧ)
'*             ：超小形クロスローラ平行ハンド(ＢＳＡ２)
'*             ：平行ハンド(ＨＡＰ)
'*             ：支点ハンド(ＨＢＬ)
'*             ：横形平行ハンド(ＨＣＰ)
'*             ：広角ハンド(ＨＤＬ)
'*             ：ベアリング平行ハンド(ＨＥＰ)
'*             ：カニ形平行ハンド(ＨＦＰ)
'*             ：ロングストローク平行ハンド(ＨＧＰ)
'*             ：トグルハンド(ＨＪＬ)
'*             ：クロスローラ平行ハンド(ＨＫＰ)
'*             ：薄形平行ハンド(ブッシュタイプ)(ＨＬＡ)
'*             ：ゴムカバー付薄形平行ハンド(ブッシュタイプ)(ＨＬＡＧ)
'*             ：薄形平行ハンド(ベアリングタイプ)(ＨＬＢ)
'*             ：ゴムカバー付薄形平行ハンド(ベアリングタイプ)(ＨＬＢＧ)
'*             ：薄形ロングストローク平行ハンド(ＨＬＣ)
'*             ：小形カニ形平行ハンド(ＨＭＦ)
'*             ：ＬＭガイド付大形カニ形平行ハンド(ＨＭＦＢ)
'*             ：３方爪ロングストロークチャック(ＣＫ)
'*             ：３方爪薄形チャック(ＣＫＡ)
'*             ：中空チャック(ＣＫＦ)
'*             ：３方爪ベアリングチャック(ＣＫＧ)
'*             ：超ロングストロークチャック(ＣＫＪ)
'*             ：高把持形広角ハンド (ＨＪＤ)
'*
'*  更新履歴
'*                                      更新日：2008/07/20      更新者：T.Sato
'*  ・受付No：RM0806061　上記コメントに「HJD」追加　※ロジック変更はなし
'*  ・受付No：RM0908030  二次電池対応機器　HMD
'*                                      更新日：2009/10/19   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'*  ・受付No：RM1001045  二次電池対応機器　BHA
'*                                      更新日：2010/02/23   更新者：Y.Miura
'*
'************************************************************************************
Module KHPrice69

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim strOptionKataban As String = ""
        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim strOptionP4 As String = String.Empty            'RM0908030 2009/10/19 Y.Miura　二次電池対応
        Dim bolC5Flag As Boolean                            'RM0908030 2009/10/19 Y.Miura　二次電池対応
        Dim strOptionFP1 As String = String.Empty

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)            'RM0908030 2009/10/28 Y.Miura

            'C5チェック                      'RM0908030 2009/10/19 Y.Miura
            'RM1001043 2010/02/22 Y.Miura 二次電池C5加算廃止
            'bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc,false)
            bolC5Flag = False

            'P4判定                        'RM0908030 2009/10/19 Y.Miura   FP1判定追加
            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "BHA"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "P" Then
                        strOptionFP1 = objKtbnStrc.strcSelection.strOpSymbol(7)
                    Else
                        If objKtbnStrc.strcSelection.strOpSymbol.Length >= 8 Then
                            strOptionP4 = objKtbnStrc.strcSelection.strOpSymbol(7)
                        End If
                    End If
                Case "HMF"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "F" Then
                        strOptionFP1 = objKtbnStrc.strcSelection.strOpSymbol(7)
                    Else
                        If objKtbnStrc.strcSelection.strOpSymbol.Length >= 8 Then
                            strOptionP4 = objKtbnStrc.strcSelection.strOpSymbol(7)
                        End If
                    End If
                Case "CKL2"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "5" Then
                        strOptionFP1 = objKtbnStrc.strcSelection.strOpSymbol(7)
                    Else
                        If objKtbnStrc.strcSelection.strOpSymbol.Length >= 8 Then
                            strOptionP4 = objKtbnStrc.strcSelection.strOpSymbol(7)
                        End If
                    End If
                Case "CKLG2"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "F" Then
                        strOptionFP1 = objKtbnStrc.strcSelection.strOpSymbol(7)
                    Else
                        If objKtbnStrc.strcSelection.strOpSymbol.Length >= 8 Then
                            strOptionP4 = objKtbnStrc.strcSelection.strOpSymbol(7)
                        End If
                    End If
                Case "BHG"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "F" Then
                        strOptionFP1 = objKtbnStrc.strcSelection.strOpSymbol(7)
                    Else
                        If objKtbnStrc.strcSelection.strOpSymbol.Length >= 8 Then
                            strOptionP4 = objKtbnStrc.strcSelection.strOpSymbol(7)
                        End If
                    End If
                Case "CKG"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "F" Then
                        strOptionFP1 = objKtbnStrc.strcSelection.strOpSymbol(7)
                    Else
                        If objKtbnStrc.strcSelection.strOpSymbol.Length >= 8 Then
                            strOptionP4 = objKtbnStrc.strcSelection.strOpSymbol(7)
                        End If
                    End If
                Case Else
                    If objKtbnStrc.strcSelection.strOpSymbol.Length >= 8 Then
                        strOptionP4 = objKtbnStrc.strcSelection.strOpSymbol(7)
                    End If
            End Select

            'オプション形番設定
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(2), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        strOptionKataban = strOptionKataban & strOpArray(intLoopCnt).Trim
                End Select
            Next

            '↓RM1310067 2013/10/23
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "HHC", "HHD", "CKT", "CKU", "HLF"
                    'C5チェック
                    bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    'If bolC5Flag = True Then
                    '    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    'End If

                    'スイッチ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(4).Trim)

                        'リード線長さ加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                        End If
                    End If
                    '↑RM1310067 2013/10/23
                Case Else
                    '基本価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0908030 2009/10/28 Y.Miura
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   strOptionKataban
                        decOpAmount(UBound(decOpAmount)) = 1
                        'RM0908030 2009/10/28 Y.Miura
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    End If

                    '小爪加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "C" Then
                            decOpAmount(UBound(decOpAmount)) = 3
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                        'RM0908030 2009/10/28 Y.Miura
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    End If

                    'スイッチ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)

                        'リード線長さ加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                        End If

                        'スイッチP4加算                          'RM0908030 2009/10/19 Y.Miura
                        If strOptionP4 <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SW-P4"
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                        End If

                        'スイッチ取付金具加算価格キー
                        Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                            Case "CKG"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "50CS"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "T"
                                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                                        'RM0908030 2009/10/28 Y.Miura
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                            Case "CKA"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "50CS", "60CS", "70CS"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "T"
                                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                                        'RM0908030 2009/10/28 Y.Miura
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                            Case "CKF"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "30CS", "40CS"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "T"
                                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                                        'RM0908030 2009/10/28 Y.Miura
                                        If bolC5Flag = True Then
                                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                        End If
                                End Select
                        End Select
                    End If

                    '食品製造工程向け商品
                    If strOptionFP1 <> "" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2)
                            Case ""
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim()
                                decOpAmount(UBound(decOpAmount)) = 1
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim() & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim()
                                decOpAmount(UBound(decOpAmount)) = 1
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End Select
                    End If

                    '本体のＰ４※加算                          'RM0908030 2009/10/19 Y.Miura
                    If strOptionP4 <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)     'RM0908030 2009/10/28 Y.Miura
                        Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                            Case "BHE"
                                If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-OP-" & strOptionP4
                                Else
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "-OP-" & strOptionP4
                                End If
                            Case "BHA"              'RM1001045 2010/02/23 Y.Miura 追加
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & strOptionP4
                            Case "HMF"
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                          objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "-OP-" & strOptionP4
                            Case Else
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                         strOptionP4 & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        End Select
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
