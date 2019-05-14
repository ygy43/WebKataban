'************************************************************************************
'*  ProgramID  ：KHPrice33
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/21   作成者：NII K.Sudoh
'*
'*  概要       ：セルシリンダ　ＣＡＶ２／ＣＯＶＰ２／ＣＯＶＮ２
'*
'*  更新履歴   ：                       更新日：2007/05/16   更新者：NII A.Takahashi
'*               ・T2W/T3Wスイッチ追加に伴い、リード線加算ロジック部を修正
'************************************************************************************
Module KHPrice33

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'ストローク取得
            intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim))

            '基本価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "N", "NS"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & CdCst.Sign.Hypen & "N" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '支持形式加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "CA", "TC", "TF"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "C*V2" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '2014/10/22 電圧加算価格キー
            If Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 1) = "A" Then
                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "DC24V" Then
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                    Dim cntZ As Integer = 0
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        If strOpArray(intLoopCnt).Trim = "Z" Then
                            cntZ = 1
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "CAV2" & CdCst.Sign.Hypen & _
                                                                        strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & "DC24V"
                            decOpAmount(UBound(decOpAmount)) = 1
                            Exit For
                        End If
                    Next
                    If cntZ <= 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "CAV2" & CdCst.Sign.Hypen & "DC24V"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            Else
                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "DC24V" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "COVP2" & CdCst.Sign.Hypen & "DC24V"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            'スイッチ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                    Case "A", "B"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "C*V2" & objKtbnStrc.strcSelection.strOpSymbol(7).Trim & _
                                                                            objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "C*V2" & objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)

                        'リード線長さ加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                Case "T1H", "T1V", "T2H", "T2V", "T2YH", "T2YV", "T3H", "T3V", _
                                     "T3YH", "T3YV", "T0H", "T0V", "T5H", "T5V", "T8H", "T8V", "T2JH", "T2JV", _
                                     "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "C*V2" & CdCst.Sign.Hypen & "SWLW(2)" & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", "T2YMV", _
                                     "T3YMH", "T3YMV"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "C*V2" & CdCst.Sign.Hypen & "SWLW(3)" & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "C*V2" & objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                            End Select
                        End If
                End Select
            End If

            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "J"
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "100" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "C*V2" & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & "75" & CdCst.Sign.Hypen & _
                                                                       intStroke.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "C*V2" & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                       intStroke.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case "TB1", "TB2", "MF1", "Z", "Q"
                        If Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 1) = "A" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "CAV2" & strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "COVP2" & strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                End Select
            Next

            '付属品加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(11), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "100" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "C*V2" & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & "75"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "C*V2" & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                End Select
            Next

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
