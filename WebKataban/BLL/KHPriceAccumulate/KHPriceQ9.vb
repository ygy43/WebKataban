'************************************************************************************
'*  ProgramID  ：KHPriceQ9
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2009/12/18   作成者：Y.Miura
'*                                      更新日：             更新者：
'*
'*  概要       ：EXAシリーズ  (圧縮空気用　パイロット式２ポート電磁弁 小形エアブローバルブ)
'*             ：GEXAシリーズ (圧縮空気用　パイロット式２ポート電磁弁 マニホールド)
'*
'************************************************************************************
Module KHPriceQ9

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            If objKtbnStrc.strcSelection.strSeriesKataban.Trim <> "GEXA" Then
                '基本価格キー
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                'シール材質加算
                If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "0" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2)
                    decOpAmount(UBound(decOpAmount)) = 1

                End If

                'コイルオプション加算
                'ケーブル長さ
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    'RM1612033 製品追加対応のため、case条件に「3」を追加  2016/12/19 追加 松原
                    Case "1", "F", "3"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Case ""
                            Case "2C"
                                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim.Equals("1") Then
                                    'AC100Vのみ加算
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                               CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    'DC24V,DC12V は加算無し
                                End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3)
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select

                        'その他オプション加算
                        If Not objKtbnStrc.strcSelection.strOpSymbol(4).Trim.Equals("") Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case "2"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Case ""
                            Case "2C"
                                If objKtbnStrc.strcSelection.strOpSymbol(4).Trim.Equals("1") Then
                                    'AC100Vのみ加算
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                               CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    'DC24V,DC12V は加算無し
                                End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3)
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select

                ''その他オプション加算
                'If Not objKtbnStrc.strcSelection.strOpSymbol(4).Trim.Equals("") Then
                '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                '                                               CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                '                                               objKtbnStrc.strcSelection.strOpSymbol(4)
                '    decOpAmount(UBound(decOpAmount)) = 1
                'End If

                '食品製造工程向け
                If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6)
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

            End If

            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "GEXA" Then
                '基本価格キー
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4)

                decOpAmount(UBound(decOpAmount)) = 1

                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "2C" Then
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "1" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5)
                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(3)
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5)
                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(3)
                End If

            End If

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module
