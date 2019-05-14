'************************************************************************************
'*  ProgramID  ：KHPriceA5
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：コンパクトロータリバルブ　ＣＨＢ（Ｒ）
'*　　　　　　　：コンパクトロータリバルブ　ＣＨＢ（Ｖ）
'*　　　　　　　：コンパクトロータリバルブ　ＣＨＧ（Ｒ）
'*　　　　　　　：コンパクトロータリバルブ　ＣＨＧ（Ｖ）
'*
'*  変更 RM1004012 2010/04/23 Y.Miura
'************************************************************************************
Module KHPriceA5

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case objKtbnStrc.strcSelection.strKeyKataban
                '屋外シリーズ
                Case "X", "W"
                    '基本価格キー
                    Select Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3)
                        Case "CHB"
                            Select Case True
                                Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim = ""
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> ""
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1

                                    'その他オプション加算価格キー
                                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                               "W-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                        If objKtbnStrc.strcSelection.strOpSymbol(1) = "WV1" Then
                                            decOpAmount(UBound(decOpAmount)) = 2
                                        Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    End If
                            End Select


                        Case Else
                            Select Case True
                                Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim = ""
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> ""
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1

                                    'その他オプション加算価格キー
                                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                               "W-" & objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                        If objKtbnStrc.strcSelection.strOpSymbol(2) = "WV1" Then
                                            decOpAmount(UBound(decOpAmount)) = 2
                                        Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    End If
                            End Select
                    End Select

                        Case Else
                            '基本価格キー
                            Select Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3)
                                Case "CHB"
                                    Select Case True
                                        Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "" And objKtbnStrc.strcSelection.strOpSymbol(3).Trim = ""
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4).Trim _
                                                                                        & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "" And objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> ""
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" And objKtbnStrc.strcSelection.strOpSymbol(3).Trim = ""
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" And objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> ""
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select

                                    'コイルオプション加算価格キー
                                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4).Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If

                                    'リミットスイッチ加算価格キー
                                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4).Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If

                                    'その他オプション加算価格キー
                                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4).Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If

                                    '二次電池追加
                                    If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 8 Then
                                        If objKtbnStrc.strcSelection.strOpSymbol(8) <> "" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4).Trim & "-OP-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    End If

                                    '食品製造工程向け対応
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-FP2"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If

                                Case Else
                                    Select Case True
                                        Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "" And objKtbnStrc.strcSelection.strOpSymbol(4).Trim = ""
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "" And objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> ""
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" And objKtbnStrc.strcSelection.strOpSymbol(4).Trim = ""
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" And objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> ""
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select

                                    'コイルオプション加算価格キー
                                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If

                                    'リミットスイッチ加算価格キー
                                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If

                                    'その他オプション加算価格キー
                                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If

                                    'RM1004012 2010/04/23 Y.Miura
                                    '二次電池追加
                                    If UBound(objKtbnStrc.strcSelection.strOpSymbol) >= 9 Then
                                        If objKtbnStrc.strcSelection.strOpSymbol(9) <> "" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    End If

                                    '食品製造工程向け対応
                                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "S" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-FP2"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                            End Select

                    End Select
        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
