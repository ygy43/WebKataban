'************************************************************************************
'*  ProgramID  ：KHPrice88
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/22   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：真空エジェクタユニット単体
'*             ：真空切替ユニット単体
'*
'************************************************************************************
Module KHPrice88

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            ' 機種毎に価格キーを設定する
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "VSK"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "**" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "**" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "VSJ"
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   "***" & CdCst.Sign.Hypen & "**" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & "*"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   "***" & CdCst.Sign.Hypen & "**" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & "*" & CdCst.Sign.Hypen & "*"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case "VSN"
                    ' 真空センサ仕様 
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   "-**-****-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   "-**-***" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1

                Case "VSJP"
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   "****" & CdCst.Sign.Hypen & "*"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   "****" & CdCst.Sign.Hypen & "*" & CdCst.Sign.Hypen & "*"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "VSNP"
                    ' 真空センサ仕様 
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   "-***-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   "-***-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1

                Case "VSX"
                    '基本キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               "***" & CdCst.Sign.Hypen & "**" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & "*"

                    '真空センサ仕様
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                    End If

                    '取付方法
                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    End If

                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'RM1806035_二次電池機種追加対応
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "P" Then

                        '基本キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(10).Trim

                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                Case "VSXP"
                    '基本キー
                    '基本キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "***" & CdCst.Sign.Hypen & "*"

                    '真空センサ仕様
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    End If

                    '取付方法
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    End If

                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "VSQ"
                    ' 基本キー
                    Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1)
                        Case "T"
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "T**" & CdCst.Sign.Hypen & "**" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "*"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "T**" & CdCst.Sign.Hypen & "**" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "*" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "D"
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "D**" & CdCst.Sign.Hypen & "**" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "*"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "D**" & CdCst.Sign.Hypen & "**" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "*" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case Else
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "**" & CdCst.Sign.Hypen & "**" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "*"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           "**" & CdCst.Sign.Hypen & "**" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & "*" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
                Case "VSQP"
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   "***" & CdCst.Sign.Hypen & "*"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   "***" & CdCst.Sign.Hypen & "*" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    'RM1806035_二次電池機種追加対応
                Case "VSFU"

                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "P" Then

                        '基本価格
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        'Ｐ４加算価格
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
