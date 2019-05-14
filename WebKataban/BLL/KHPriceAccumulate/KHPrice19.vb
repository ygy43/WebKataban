'************************************************************************************
'*  ProgramID  ：KHPrice19
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＰＫＡ／ＰＫＳ／ＰＫＷ／ＰＶＳ
'*
'************************************************************************************
Module KHPrice19

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim strStdVoltageFlag As String = CdCst.VoltageDiv.Standard
        Dim strOp As String = ""

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '価格キー設定
            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 1) = "K" Then
                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "C" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            Else
                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "4" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "NO" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            End If

            'オプション価格
            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "PKS" Then
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
            Else
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
            End If
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '電圧加算
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "PKA", "PVS"
                    strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "M" Then
                        strOp = Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2) & "M"
                    Else
                        strOp = Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2)
                    End If
                Case "PKS"
                    strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                    strOp = Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2)
                Case "PKW"
                    strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "M" Then
                        strOp = Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2) & "M"
                    Else
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            Case "AC100V", "AC200V"
                                strOp = Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2) & "M"
                            Case Else
                                strOp = Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2)
                        End Select
                    End If
            End Select
            If strStdVoltageFlag = CdCst.VoltageDiv.Standard Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                           "STD" & strOp
                decOpAmount(UBound(decOpAmount)) = 1
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                           "OTH" & strOp
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
