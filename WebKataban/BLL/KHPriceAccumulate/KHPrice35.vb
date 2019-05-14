'************************************************************************************
'*  ProgramID  ：KHPrice35
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＡＢ（防爆）／ＡＧ（防爆）
'*
'************************************************************************************
Module KHPrice35

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strCountryCd As String = Nothing, _
                                   Optional ByRef strOfficeCd As String = Nothing)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 2) = "EX" Then
                Select Case True
                    Case Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "H"
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "5" Or Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "4" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & "03"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & "0" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "J"
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "5" Or Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "4" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & "B3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & "B" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "K"
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "5" Or Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "4" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & "C3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & "C" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "L"
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "5" Or Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "4" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & "D3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & "D" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "M"
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "5" Or Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "4" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & "E3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & "E" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "N"
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "5" Or Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "4" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & "F3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & "F" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Else
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "5" Or Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "4" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                End Select
            Else
                Select Case True
                    Case Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "H"
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "5" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & "03"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & "0" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "J"
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "5" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & "B3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & "B" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "K"
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "5" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & "C3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & "C" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "L"
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "5" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & "D3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & "D" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "M"
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "5" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & "E3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & "E" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1) = "N"
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "5" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & "F3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & "F" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case Else
                        If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) = "5" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "3"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                End Select
            End If
            '支持形式加算価格キー
            Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2)
                Case "CA", "TC", "TF"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "C*V2" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'コイル別電圧加算価格キー
            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 2) = "EX" Then
                If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) <> "3" And _
                 Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) <> "4" Or _
                 objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> CdCst.PowerSupply.Const1 And _
                 objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> CdCst.PowerSupply.Const2 Then
                    If objKtbnStrc.strcSelection.strSeriesKataban.Trim <> "AG41E4" Then
                        If KHKataban.fncVoltageIsStandard(objKtbnStrc.strcSelection.strOpSymbol(8).Trim, strCountryCd, strOfficeCd) Then
                            '異電圧
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim & _
                                                                                Left(objKtbnStrc.strcSelection.strOpSymbol(8).Trim, 2)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Else
                        'コイルオプション加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim & _
                                                                            Left(objKtbnStrc.strcSelection.strOpSymbol(8).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            Else
                If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) <> "3" And _
                   Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) <> "4" Or _
                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> CdCst.PowerSupply.Const1 And _
                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> CdCst.PowerSupply.Const2 Then
                    If objKtbnStrc.strcSelection.strSeriesKataban.Trim <> "AG41E4" Then
                        '2010/08/26 ADD RM0808112(異電圧対応) START--->

                        If KHKataban.fncVoltageIsStandard(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, strCountryCd, strOfficeCd) Then
                            '異電圧
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim & _
                                                                                Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 2)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        'ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        'strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim & _
                        '                                                    Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 2)
                        'decOpAmount(UBound(decOpAmount)) = 1
                        '2010/08/26 ADD RM0808112(異電圧対応) <--- END
                    Else
                        'コイルオプション加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim & _
                                                                            Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            End If
         
            '外部導線引込方式加算価格キー
            Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 1)
                Case "L", "M", "N", "P"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション・付属品価格
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "A*4*" & strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
