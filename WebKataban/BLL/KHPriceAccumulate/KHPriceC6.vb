'************************************************************************************
'*  ProgramID  ：KHPriceC6
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ウィータハンマ緩和形電磁弁　ＷＨＬ１１
'*             ：圧縮空気用パイロット電磁弁　ＦＡＤ
'*             ：水用小形パイロット式電磁弁　ＦＷＤ
'*
'************************************************************************************
Module KHPriceC6

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "WHL11", "FAD"
                    '基本価格
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               "2C"
                    decOpAmount(UBound(decOpAmount)) = 1

                    'コイルオプション加算価格
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "2C" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If


                    'その他オプション加算価格（手動操作／取付板）
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "FWD11"
                    '基本価格
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                    'If (objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "10" Or objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "15") And _
                    '   (objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "A") And (objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "0") And _
                    '   (objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "2C") And (objKtbnStrc.strcSelection.strOpSymbol(5).Trim = " ") Then

                    'End If

            End Select

            ''コイルオプション加算価格
            'If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "2C" Then
            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
            '                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
            '    decOpAmount(UBound(decOpAmount)) = 1
            'End If


            ''その他オプション加算価格（手動操作／取付板）
            'If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
            '                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
            '    decOpAmount(UBound(decOpAmount)) = 1
            'End If

            'Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
            '    Case "FWD11"
            '        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
            '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
            '                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
            '            decOpAmount(UBound(decOpAmount)) = 1

            '        End If

            '        If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
            '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
            '                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
            '            decOpAmount(UBound(decOpAmount)) = 1

            '        End If

            'End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
