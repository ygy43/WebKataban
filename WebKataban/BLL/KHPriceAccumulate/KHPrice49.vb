'************************************************************************************
'*  ProgramID  ：KHPrice49
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/25   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ジャスフィットバルブ
'*             ：Ｆ＊Ｂ／Ｆ＊Ｇ／ＧＦ＊Ｂ／ＧＦ＊Ｇ
'*
'************************************************************************************
Module KHPrice49

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strStdVoltageFlag As String
        Dim strSeriesKataban As String
        Dim intValveQty As Integer
        Dim intMaskingQty As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'シリーズ形番設定
            If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "F" Then
                strSeriesKataban = objKtbnStrc.strcSelection.strSeriesKataban.Trim
            Else
                strSeriesKataban = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 3)
            End If

            '電磁弁＆マスキングプレート数設定
            If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "G" Then
                If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "X" Then
                    intValveQty = CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                    intMaskingQty = CInt(objKtbnStrc.strcSelection.strOpSymbol(10).Trim)
                Else
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) = 0 Then
                        intValveQty = 1
                    Else
                        intValveQty = CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                    End If
                    intMaskingQty = 0
                End If
            Else
                intValveQty = 1
                intMaskingQty = 0
            End If

            '基本価格キー
            If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "F" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                decOpAmount(UBound(decOpAmount)) = intValveQty
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                decOpAmount(UBound(decOpAmount)) = intValveQty
            End If

            'コイルオプション加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                Case "2G", "4A"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strSeriesKataban & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    decOpAmount(UBound(decOpAmount)) = intValveQty
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strSeriesKataban & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    decOpAmount(UBound(decOpAmount)) = intValveQty
            End Select

            '手動装置加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "A" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strSeriesKataban & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    decOpAmount(UBound(decOpAmount)) = intValveQty
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strSeriesKataban & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    decOpAmount(UBound(decOpAmount)) = intValveQty
                End If
            End If

            'その他オプション加算価格キー
            If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) <> "G" Then
                If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = strSeriesKataban & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                    decOpAmount(UBound(decOpAmount)) = intValveQty
                End If
            End If

            '電圧加算価格キー
            If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "G" Then
                strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                Select Case strStdVoltageFlag
                    Case CdCst.VoltageDiv.Standard
                    Case CdCst.VoltageDiv.Options
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "OPT" & CdCst.Sign.Hypen & Left(objKtbnStrc.strcSelection.strOpSymbol(8).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case CdCst.VoltageDiv.Other
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "OTH" & CdCst.Sign.Hypen & Left(objKtbnStrc.strcSelection.strOpSymbol(8).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Else
                strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                Select Case strStdVoltageFlag
                    Case CdCst.VoltageDiv.Standard
                    Case CdCst.VoltageDiv.Options
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "OPT" & CdCst.Sign.Hypen & Left(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case CdCst.VoltageDiv.Other
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "OTH" & CdCst.Sign.Hypen & Left(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            End If

            'マスキングプレート加算価格キー
            If intMaskingQty <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "-MP-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                decOpAmount(UBound(decOpAmount)) = intMaskingQty
            End If

            'マニホールドベース加算価格キー
            If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "G" Then
                If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "0" Then
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "X" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "-BS-" & _
                                                                   (intMaskingQty + intValveQty).ToString & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "-BS-" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
