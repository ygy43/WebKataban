'************************************************************************************
'*  ProgramID  ：KHPrice87
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/21   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：スーパーマウントシリンダ　ＳＭＤ２／ＳＭＤ２－Ｌ
'*
'*  ・受付No：RM0908030  二次電池対応機器　
'*                                      更新日：2009/09/04   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'*  ・受付No：RM1112XXX  SMGシリーズ追加　
'*                                      更新日：2011/12/22   更新者：Y.Tachi
'************************************************************************************
Module KHPrice87

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer
        Dim bolOptionP4 As Boolean = False      'RM0908030 2009/09/04 Y.Miura　二次電池対応
        Dim bolC5Flag As Boolean

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)                        'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応

            'RM0908030 2009/09/04 Y.Miura　二次電池対応
            Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                Case "P4", "P40"
                    bolOptionP4 = True
            End Select

            'C5チェック
            'RM1001043 2010/02/22 Y.Miura 廃止
            'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
            'bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)
            'bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc)
            bolC5Flag = False

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "SMG"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "2"
                            bolC5Flag = True

                            '基本価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "M" Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "35" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "40"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "45" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            Else
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "35" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "40"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "45" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "55" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "60"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Else
                                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "65" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "70"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Else
                                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "75" Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "80"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                Else
                                                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "85" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "90"
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    Else
                                                        If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "95" Then
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "100"
                                                            decOpAmount(UBound(decOpAmount)) = 1
                                                        Else
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                                                                                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                                            decOpAmount(UBound(decOpAmount)) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                End If

                                    '支持形式加算
                                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "M" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If

                                    '微速F加算
                                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            Case "6", "10", "16"
                                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 15 Then
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                           "VAR" & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                                                           "5"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 30 Then
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                           "VAR" & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                                                           "16"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                End If
                                            Case "20", "25", "32"
                                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 25 Then
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                           "VAR" & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                                                           "5"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 50 Then
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                           "VAR" & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                                                           "26"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                End If
                                        End Select
                                    End If


                                    'スイッチ加算価格キー
                                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If

                                        'リード線長さ加算価格キー
                                        If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If

                                        If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "D" Then
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            Else
                                                If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "T" Then
                                                    decOpAmount(UBound(decOpAmount)) = 3
                                                Else
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                End If
                                            End If
                                        End If
                                    End If

                        Case Else

                            Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                Case "P4", "P40"
                                    bolOptionP4 = True
                            End Select

                            'スイッチがK3P＊の場合はＣ５
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "K3PH" Or _
                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "K3PV" Then
                                bolC5Flag = True
                            End If

                            'ねじがNN,GNの場合はＣ５
                            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "NN" Or _
                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "GN" Then
                                bolC5Flag = True
                            End If

                            ''微速Fの場合はＣ５
                            'If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "F" Then
                            '    bolC5Flag = True
                            'End If

                            ''クリーン仕様P5,P51,P7,P71の場合はＣ５
                            'If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "P5" Or _
                            '    objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "P51" Or _
                            '    objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "P7" Or _
                            '    objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "P71" Then
                            '    bolC5Flag = True
                            'End If

                            '基本価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then

                                '2016/12/06 問い合わせ対応（バグ修正）
                                '価格キーに使用するために、マスタよりストロークを取得
                                intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                                      CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim), _
                                                                      CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim))
                                '2016/12/06 修正End

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                                '2016/12/06 問い合わせ対応（バグ修正）
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                           intStroke.ToString

                                'strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                '                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                '                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                '                                           objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                '2016/12/06 修正End
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "35" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "40"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "45" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "50"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "55" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "60"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Else
                                            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "65" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "70"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Else
                                                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "75" Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "80"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                Else
                                                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "85" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "90"
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    Else
                                                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "95" Then
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "100"
                                                            decOpAmount(UBound(decOpAmount)) = 1
                                                        Else
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                                                       objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                                            decOpAmount(UBound(decOpAmount)) = 1
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If

                            '微速F加算
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "6", "10", "16"
                                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <= 15 Then
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                   "VAR" & CdCst.Sign.Hypen & _
                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                   "5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(6).Trim <= 30 Then
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                   "VAR" & CdCst.Sign.Hypen & _
                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                   "16"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    Case "20", "25", "32"
                                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <= 25 Then
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                   "VAR" & CdCst.Sign.Hypen & _
                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                   "5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(6).Trim <= 50 Then
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                   "VAR" & CdCst.Sign.Hypen & _
                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                   "26"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                End Select
                            End If


                            'スイッチ加算価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If

                                'リード線長さ加算価格キー
                                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If

                                If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "D" Then
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Else
                                        If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "T" Then
                                            decOpAmount(UBound(decOpAmount)) = 3
                                        Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    End If
                                End If

                                If bolOptionP4 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SW-P4"
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                End If
                            End If

                            'クリーン仕様加算
                            If objKtbnStrc.strcSelection.strOpSymbol(10).Trim <> "" Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                    Case "P5", "P51"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                                   "P5"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case "P7", "P71"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                                   "P7"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case "P4"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                                   "P4"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case "P40"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                                   "P40"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            End If
                    End Select
                Case Else

                    'ストローク取得
                    intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                          CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                          CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim))

                    'バリエーション(微速)加算価格キー
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "F"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "6", "10", "16"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 15
                                            'ストローク5～15
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                                       CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR5" & CdCst.Sign.Hypen & "15"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 16 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 30
                                            'ストローク16～30
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                                       CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR16" & CdCst.Sign.Hypen & "30"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 31
                                            'ストローク31～60
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                                       CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR31" & CdCst.Sign.Hypen & "60"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                                Case "20", "25", "32"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 25
                                            'ストローク5～25
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                                       CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR5" & CdCst.Sign.Hypen & "25"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 26 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 50
                                            'ストローク26～50
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                                       CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR26" & CdCst.Sign.Hypen & "50"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 51
                                            'ストローク51～100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                                       CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR51" & CdCst.Sign.Hypen & "100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                            If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select
                            End Select
                    End Select

                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If

                    'マグネット内臓(L)加算価格キー
                    If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6, 1) = "L" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                   CdCst.Sign.Hypen & "L" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    End If

                    '支持形式加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "DC" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then    'RM0912XXX 2009/12/09 Y.Miura　二次電池C5加算対応
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    End If

                    'スイッチ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)

                        'リード線長さ加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "K0H", "K0V", "K2H", "K2V", "K3H", _
                                     "K3V", "K5H", "K5V", "K2YH", "K2YV", _
                                     "K3YH", "K3YV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                                Case "K2YFH", "K2YFV", "K3YFH", "K3YFV", "K2YMH", _
                                     "K2YMV", "K3YMH", "K3YMV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim & "Y"
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                            End Select
                        End If

                        'RM0908030 2009/09/04 Y.Miura　二次電池対応
                        'Ｐ４加算　SW数
                        If bolOptionP4 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-SW-P4"
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                        End If

                    End If

                    'クリーン仕様加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
