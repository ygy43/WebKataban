'************************************************************************************
'*  ProgramID  ：KHPriceK9
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/25   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ガス燃焼システム
'*             ：ＧＡＶ
'*             ：ＤＳＣ
'*             ：ＤＳＧ－Ｗ
'*             ：ＶＮＡ
'*             ：ＶＮＡ－Ｒ・ＲＨ
'*             ：ＭＮ
'*             ：ＶＬＡ
'*             ：ＶＮＲ
'*             ：ＶＮＭ
'*             ：ＶＬＭ
'*             ：ＨＫ（ＨＫ１）
'*             ：ＨＳ
'*             ：ＤＧ
'*             ：ＤＬ
'*
'************************************************************************************
Module KHPriceK9

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strStdVoltageFlag As String
        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim bolOptionY As Boolean
        Dim bolOptionB As Boolean
        Dim bolOptionP As Boolean
        Dim bolOptionM As Boolean                      '2011/03/07 ADD RM1103016(4月VerUP:HVB,PDV2,AB71(他)シリーズ)
        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "GAV"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'その他オプション加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case "DSG"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                        '基本価格キー
                        Select Case True
                            Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "S" Or _
                                 Right(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "S"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "S"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select

                        'オプション加算価格キー
                        bolOptionY = False
                        bolOptionB = False
                        bolOptionP = False
                        '2011/03/07 ADD RM1103016(4月VerUP:HVB,PDV2,AB71(他)シリーズ) START--->
                        bolOptionM = False
                        '2011/03/07 ADD RM1103016(4月VerUP:HVB,PDV2,AB71(他)シリーズ) <---END
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(2), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "SY", "Y", "BY", "PY"
                                    bolOptionY = True
                            End Select
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "BY", "B", "BP"
                                    bolOptionB = True
                            End Select
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "PS", "PY", "BP", "P"
                                    bolOptionP = True
                                    '2011/03/07 ADD RM1103016(4月VerUP:HVB,PDV2,AB71(他)シリーズ) START--->
                                Case "3M"
                                    bolOptionM = True
                                    '2011/03/07 ADD RM1103016(4月VerUP:HVB,PDV2,AB71(他)シリーズ) <---END
                            End Select
                        Next
                        If bolOptionY = True Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "Y"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If bolOptionB = True Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "B"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        If bolOptionP = True Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "P"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        '2011/03/07 ADD RM1103016(4月VerUP:HVB,PDV2,AB71(他)シリーズ) START--->
                        If bolOptionM = True Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "3M"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        '2011/03/07 ADD RM1103016(4月VerUP:HVB,PDV2,AB71(他)シリーズ) <---END

                        '電圧加算価格キー
                        strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim)
                        Select Case strStdVoltageFlag
                            Case CdCst.VoltageDiv.Standard
                            Case CdCst.VoltageDiv.Options
                            Case CdCst.VoltageDiv.Other
                                If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 2) = CdCst.PowerSupply.Div1 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-AC-OTH"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-DC-OTH"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select

                        'その他オプション加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Else
                        '基本価格キー
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            Case "W", "WK", "WY"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "WK"
                                decOpAmount(UBound(decOpAmount)) = 1

                                'オプション加算価格キー
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "Y"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select

                        '電圧加算価格キー
                        strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim)
                        Select Case strStdVoltageFlag
                            Case CdCst.VoltageDiv.Standard
                            Case CdCst.VoltageDiv.Options
                            Case CdCst.VoltageDiv.Other
                                If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 2) = CdCst.PowerSupply.Div1 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-AC-OTH"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-DC-OTH"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                    End If
                Case "VNA"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        '能力加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        '↓RM1311065 2013/11/21 修正
                        ''その他オプション加算価格キー
                        'If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                        '                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                        '                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        '    decOpAmount(UBound(decOpAmount)) = 1
                        'End If

                        'オプション加算価格キー
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                       CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Next
                        '↑RM1311065 2013/11/21 修正

                        '電圧加算価格キー
                        strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                        Select Case strStdVoltageFlag
                            Case CdCst.VoltageDiv.Standard
                            Case CdCst.VoltageDiv.Options
                            Case CdCst.VoltageDiv.Other
                                If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2) = CdCst.PowerSupply.Div1 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-AC-OTH-" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-DC-OTH-" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                    Else
                        '基本価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "H" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-R-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-R-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        '電圧加算価格キー
                        strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                        Select Case strStdVoltageFlag
                            Case CdCst.VoltageDiv.Standard
                            Case CdCst.VoltageDiv.Options
                            Case CdCst.VoltageDiv.Other
                                If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2) = CdCst.PowerSupply.Div1 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-AC-OTH-" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-DC-OTH-" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                    End If
                Case "MN"
                    '基本価格キー
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "W" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'オプション加算価格キー
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "PF1/2", "C19", "W-PF1/2", "W-C19"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    '電圧加算価格キー
                    strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim)
                    Select Case strStdVoltageFlag
                        Case CdCst.VoltageDiv.Standard
                        Case CdCst.VoltageDiv.Options
                        Case CdCst.VoltageDiv.Other
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 2) = CdCst.PowerSupply.Div1 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-AC-OTH"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-DC-OTH"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
                Case "VLA"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '能力加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '↓RM1311065 2013/11/21 修正
                    ''その他オプション加算価格キー
                    'If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                    '                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                    '                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    '    decOpAmount(UBound(decOpAmount)) = 1
                    'End If

                    'オプション加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                   CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Next
                    '↑RM1311065 2013/11/21 修正

                    '電圧加算価格キー
                    strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                    Select Case strStdVoltageFlag
                        Case CdCst.VoltageDiv.Standard
                        Case CdCst.VoltageDiv.Options
                        Case CdCst.VoltageDiv.Other
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2) = CdCst.PowerSupply.Div1 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-AC-OTH-" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-DC-OTH-" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
                Case "VNM"
                    '基本価格キー
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "E", "K"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    '電圧加算価格キー
                    strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim)
                    Select Case strStdVoltageFlag
                        Case CdCst.VoltageDiv.Standard
                        Case CdCst.VoltageDiv.Options
                        Case CdCst.VoltageDiv.Other
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 2) = CdCst.PowerSupply.Div1 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-AC-OTH"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-DC-OTH"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
                Case "VLM"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '電圧加算価格キー
                    strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim)
                    Select Case strStdVoltageFlag
                        Case CdCst.VoltageDiv.Standard
                        Case CdCst.VoltageDiv.Options
                        Case CdCst.VoltageDiv.Other
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 2) = CdCst.PowerSupply.Div1 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-AC-OTH"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-DC-OTH"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
                Case "HK"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '丸形端子箱加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'マイクロスイッチ種類加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '屋外仕様加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '電圧加算価格キー
                    strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                    Select Case strStdVoltageFlag
                        Case CdCst.VoltageDiv.Standard
                        Case CdCst.VoltageDiv.Options
                        Case CdCst.VoltageDiv.Other
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 2) = CdCst.PowerSupply.Div1 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-AC-OTH"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-DC-OTH"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
                Case "HS"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '丸形端子箱加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'マイクロスイッチ種類加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '屋外仕様加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '電圧加算価格キー
                    strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim)
                    Select Case strStdVoltageFlag
                        Case CdCst.VoltageDiv.Standard
                        Case CdCst.VoltageDiv.Options
                        Case CdCst.VoltageDiv.Other
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2) = CdCst.PowerSupply.Div1 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-AC-OTH"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-DC-OTH"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
                Case "DL"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '接点加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '電線管ネジソケット加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case "DG"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '接点加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '電線管ネジソケット加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case "VNR"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    '電圧加算価格キー
                    strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim)
                    Select Case strStdVoltageFlag
                        Case CdCst.VoltageDiv.Standard
                        Case CdCst.VoltageDiv.Options
                        Case CdCst.VoltageDiv.Other
                            If Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 2) = CdCst.PowerSupply.Div1 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-AC-OTH-" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-DC-OTH-" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
                Case "PXB-B"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'オプション加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "ZB4-B" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case "PXC-K"
                    '基本価格キー
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "211"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                Case "01", "02", "05", "06", "21"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & objKtbnStrc.strcSelection.strOpSymbol(1).Trim _
                                                                             & objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "ZCK-D" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Case "221"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                Case "05", "06"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & objKtbnStrc.strcSelection.strOpSymbol(1).Trim _
                                                                             & objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "ZCK-D" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select


                    End Select

                    'オプション加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "ZCK-Y" & objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
