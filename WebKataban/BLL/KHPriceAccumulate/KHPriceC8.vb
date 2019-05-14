'************************************************************************************
'*  ProgramID  ：KHPriceC8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/06   作成者：NII K.Sudoh
'*
'*  概要       ：残圧排出弁／圧力スイッチ　白色シリーズ
'*             ：Ｖ１０００－Ｗ／Ｖ３０００－Ｗ／Ｖ３０１０－Ｗ
'*　　　　　　 ：Ｐ４０００－Ｗ／Ｐ４１００－Ｗ／Ｐ８１００－Ｗ
'*　　　　　　 ：ＡＰＳ－Ｗ
'*　　　　　　 ：ＦＸ
'*
'*  更新履歴   ：                       更新日：2008/01/22   更新者：NII A.Takahashi
'*               ・APS-W/V1000-W/V3000-W/V3010-W/P8100-Wを追加したため、単価見積りロジック変更
'*  ・受付No：RM0907070  二次電池対応機器　V3010
'*                                      更新日：2009/08/25   更新者：Y.Miura
'*  ・受付No：RM1001045  二次電池対応機器追加 P4100 
'*                                      更新日：2010/02/24   更新者：Y.Miura
'************************************************************************************
Module KHPriceC8

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                'RM1306001 2013/06/04 追加
                Case "FX1004", "FX1011", "FX1037"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "" Then
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4), 1) = "F" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "F"
                        Else
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim

                        End If
                    Else
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4), 1) = "F" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "F"
                        Else
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim

                        End If
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション加算価格キー
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5)
                        Case "Z"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-Z"

                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "M"
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(4), 1) = "C" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-CM"

                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-M"

                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "M1"
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(4), 1) = "C" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-CM1"

                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-M1"

                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case Else
                    End Select

                    'RM1807033_FP機種追加、二次電池機種価格不具合修正
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "4", "F"

                            '二次電池、食品製造加算価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim

                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            'アタッチメント加算価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim

                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                        Case Else

                            'アタッチメント加算価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim

                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                    End Select

                Case "P4000"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)

                    'RM1612036 T8オプション追加による変更  2016/12/19 追加 松原
                    If objKtbnStrc.strcSelection.strOpSymbol(3).IndexOf("T8") >= 0 Then
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim

                    ElseIf objKtbnStrc.strcSelection.strOpSymbol(3).IndexOf("T") >= 0 Then
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & "T"
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "T"
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    'アタッチメント加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                Case "V1000", "V3000", "V3010", "V6010"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'アタッチメント加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(5), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    '食品製造工程向け商品
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4)
                        Case "FP1"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    '食品製造工程向け商品
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "V1000", "V3000"
                            If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
                                For intLoopCnt = 0 To strOpArray.Length - 1
                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case ""
                                        Case Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                                       strOpArray(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Next
                            End If
                        Case Else
                    End Select

                    'RM0907070 2009/08/25 Y.Miura　二次電池対応
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                    'RM1001045 2010/02/24 Y.Miura 二次電池追加
                    '2011/10/24 ADD RM1110032(11月VerUP:二次電池) START--->
                    '二次電池用
                    If objKtbnStrc.strcSelection.strKeyKataban = "X" Then
                        '二次電池加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                        End If
                    End If
                    '2011/10/24 ADD RM1110032(11月VerUP:二次電池) <---END
                    'Case "P1100", "P4100", "P8100"
                Case "P1100", "P8100"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strKeyKataban.Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    ''添付品価格キー
                    'If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) & CdCst.Sign.Asterisk & _
                    '                                               Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 5) & _
                    '                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    '    decOpAmount(UBound(decOpAmount)) = 1
                    'End If

                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "F"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "P*100" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'RM1305XXX 2013/05/02
                            '添付品価格キー
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(5), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "P*100" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Next

                            'リード線価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                            'RM1001045 2010/02/24 Y.Miura 二次電池追加

                        Case Else
                            'RM1305XXX 2013/05/02
                            '添付品価格キー
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "P*100" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Next

                            'リード線価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                            'RM1001045 2010/02/24 Y.Miura 二次電池追加

                    End Select


                Case "P4100"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strKeyKataban.Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション加算（二次電池）
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "P4"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-OP-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'オプション加算（食品製造向け商品）
                        Case "FP1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "P*100" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1


                    End Select
                    ''添付品価格キー
                    'If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) & CdCst.Sign.Asterisk & _
                    '                                               Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 5) & _
                    '                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    '    decOpAmount(UBound(decOpAmount)) = 1
                    'End If

                    'RM1305XXX 2013/05/02
                    '添付品価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(5), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "P*100" & strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Next

                    'リード線価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "APS"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '2011/10/24 ADD RM1110032(11月VerUP:二次電池) START--->
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    'strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                    '                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                    '                                           objKtbnStrc.strcSelection.strKeyKataban.Trim
                    '2011/10/24 ADD RM1110032(11月VerUP:二次電池) <---END
                    decOpAmount(UBound(decOpAmount)) = 1

                    'リード線価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '添付品加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    '2011/10/24 ADD RM1110032(11月VerUP:二次電池) START--->
                    '二次電池用
                    If objKtbnStrc.strcSelection.strKeyKataban = "Z" Then
                        '二次電池加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                        End If
                    End If
                    '2011/10/24 ADD RM1110032(11月VerUP:二次電池) <---END
                Case "P1100-UN", "P4100-UN"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 8) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "-W"
                    decOpAmount(UBound(decOpAmount)) = 1

                    'リード線価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
