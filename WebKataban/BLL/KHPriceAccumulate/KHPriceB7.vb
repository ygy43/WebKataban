'************************************************************************************
'*  ProgramID  ：KHPriceB7
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/05   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：真空エジェクタユニットマニホールド(ベース／単体)
'*             ：真空切替ユニットマニホールド(ベース／単体)
'*
'************************************************************************************
Module KHPriceB7

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '機種毎に価格キーを設定
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "VSKM"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "1"
                            'VSKM-**A
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-**" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim

                            '真空ポート
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "T" Then
                                'VSKM-**A-T*
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-T*"
                            Else
                                'VSKM-**A-*"
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-*"
                            End If
                            '電磁弁電圧
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                'VSKM-**A-*"
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-*"

                                'バルブタイプ
                                If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                                    'VSKM-**A-*"
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "*"
                                End If
                            Else
                                'バルブタイプ
                                If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                                    'VSKM-**A-*"
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-*"
                                End If
                            End If
                            '真空センサ仕様
                            If objKtbnStrc.strcSelection.strOpSymbol(10).Trim <> "" Then
                                'VSKM-**A-*"
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-*"
                            End If

                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2"
                            'VSKM-***-2
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-***-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case "VSXM"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "1"
                            'VSXM-***-*-*
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-***-*-*"

                            '真空センサ仕様
                            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                                'VSXM-***-*-*
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            End If

                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2"
                            'VSXM-**-2
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-**-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case "VSZM"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "1"
                            'VSZM-**-*
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-**-*"

                            '真空センサ仕様
                            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                                'VSZM-**-*-DA
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            End If

                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2"
                            'VSZM-V*-3
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "*-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "3"
                            'VSZM-**-2-**
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-**-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim & "-**"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case "VSXPM"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "1"
                            'VSXPM-D*-*
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "*-*"

                            '真空センサ仕様
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                                'VSXPM-***-*-*
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            End If

                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2"
                            'VSXM-**-2
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "VSXM-**-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case "VSZPM"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "1"
                            'VSZPM-*
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-*"

                            '真空センサ仕様
                            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                                'VSZPM-*-DW
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            End If

                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2"
                            'VSZPM-V-3
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-V-3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "3"
                            'VSZPM-**-2-**
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-**-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim & "-**"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
