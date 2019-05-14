'************************************************************************************
'*  ProgramID  ：KHPrice06
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/08   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：セレックスバルブ
'*             ：３ＫＡ１
'*             ：４ＫＡ１／２／３／４
'*             ：４ＫＢ１／２／３／４
'*             ：Ｍ３ＫＡ１／２／３／４
'*             ：Ｍ４ＫＡ１／２／３／４
'*             ：Ｍ４ＫＢ１／２／３／４
'*             ：４ＨＡ１／２
'*             ：Ｍ４ＨＡ１／２
'************************************************************************************
Module KHPrice06

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strCountryCd As String = Nothing, _
                                   Optional ByRef strOfficeCd As String = Nothing)



        Dim strStdVoltageFlag As String = CdCst.VoltageDiv.Standard
        Dim intLoopCnt As Integer
        Dim intLoopCnt1 As Integer

        Dim strOpKtbn As String = Nothing
        Dim intIndex As Integer = 0
        Dim intStationQty As Integer = 0
        Dim intQuantity As Integer = 0
        Dim intQuantity1 As Integer = 0
        Dim intQuantity2 As Integer = 0
        Dim strOpPort As String = Nothing
        Dim bolSkipFlg As Boolean = False
        Dim kata As String

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '要素取得
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                'RM1805001_4Rシリーズ追加
                Case "M4HA1", "M4HA2", "M4HA3", "4HA1", "4HA2", "4HA3", _
                     "M4JA1", "M4JA2", "M4JA3", "4JA1", "4JA2", "4JA3", _
                     "M4RD1", "M4RD2", "M4RE1", "M4RE2", "4RD1", "4RD2", "4RE1", "4RE2"
                Case Else
                    Call subElementChk(objKtbnStrc, intIndex, strOpPort)
            End Select

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "M4HA1", "M4HA2", "M4HA3", "4HA1", "4HA2", "4HA3"
                    '基本価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                        '仕様有り
                        For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                            If Len(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim) <> 0 And _
                               objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                                If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2) = "MP" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "C"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                            End If
                        Next

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        'Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                        '    Case "M4HA1", "M4HA2"
                        '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & _
                        '                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-" & _
                        '                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        '        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                        '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        '        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        '        decOpAmount(UBound(decOpAmount)) = 1
                        '    Case "4HA1", "4HA2"
                        '仕様なし
                        If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                            kata = objKtbnStrc.strcSelection.strSeriesKataban.Trim
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "9-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(4).Trim

                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        'End Select
                    End If
                Case "M4JA1", "M4JA2", "M4JA3", "4JA1", "4JA2", "4JA3"
                    '基本価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                        '仕様有り
                        For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                            If Len(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim) <> 0 And _
                               objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                                If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2) = "MP" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "C"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                            End If
                        Next

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        'オプション加算価格キー
                        For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                            If Len(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim) <> 0 And _
                               objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                                If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2) = "MP" Then

                                ElseIf Right(Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5), 1) = "1" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1 * objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    decOpAmount(UBound(decOpAmount)) = 2 * objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                            End If
                        Next

                    Else
                        '仕様なし
                        If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                            kata = objKtbnStrc.strcSelection.strSeriesKataban.Trim
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "9-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(5).Trim

                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                        'End Select
                        'オプション加算価格キー
                        If Not objKtbnStrc.strcSelection.strOpSymbol(4).Length = 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            If Left(objKtbnStrc.strcSelection.strSeriesKataban, 1) = "M" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Then
                                    decOpAmount(UBound(decOpAmount)) = 1 * objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 2 * objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                End If
                            Else
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Then
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 2
                                End If
                            End If
                        End If
                    End If
                    ''RM1805001_4Rシリーズ追加
                Case "M4RD1", "M4RD2", "M4RE1", "M4RE2", "4RD1", "4RD2", "4RE1", "4RE2"
                    '基本価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                        '仕様有り
                        For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                            If Len(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim) <> 0 And _
                               objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                                If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2) = "MP" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim,2) & "C"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                        Case "M4RD1", "M4RD2"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            'GS
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) & "-" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Case Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-00"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End Select
                                End If
                            End If
                        Next

                        '連数
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        'オプション加算価格キー
                        For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                            If Len(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim) <> 0 And _
                               objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                                If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2) = "MP" Then

                                ElseIf Right(Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5), 1) = "1" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1 * objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    decOpAmount(UBound(decOpAmount)) = 2 * objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                            End If
                        Next


                        For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                            If Len(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim) <> 0 And _
                               objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                                If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2) = "MP" Then

                                ElseIf Right(Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5), 1) = "1" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1 * objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "M1" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1 * objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    decOpAmount(UBound(decOpAmount)) = 2 * objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                            End If
                        Next

                    Else
                        '仕様なし
                        If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then

                            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                Case "M4RD1", "M4RD2"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "9"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(6).Trim

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "9-00"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            End Select

                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                Case "4RD1", "4RD2"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        End If
                        'End Select
                        'オプション加算価格キー
                        If Not objKtbnStrc.strcSelection.strOpSymbol(4).Length = 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            If Left(objKtbnStrc.strcSelection.strSeriesKataban, 1) = "M" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Then
                                    decOpAmount(UBound(decOpAmount)) = 1 * objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 2 * objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                End If
                            Else
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Then
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 2
                                End If
                            End If
                        End If

                        'オプション加算価格キー
                        If Not objKtbnStrc.strcSelection.strOpSymbol(5).Length = 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            If Left(objKtbnStrc.strcSelection.strSeriesKataban, 1) = "M" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & "-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Then
                                    decOpAmount(UBound(decOpAmount)) = 1 * objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "M1" Then
                                    decOpAmount(UBound(decOpAmount)) = 1 * objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 2 * objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                End If
                            Else
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Then
                                    decOpAmount(UBound(decOpAmount)) = 1
                                ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "M1" Then
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 2
                                End If
                            End If
                        End If
                    End If
                Case Else
                    '基本価格キー
                    If objKtbnStrc.strcSelection.strSpecNo.Trim <> "" And _
                       (objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "80" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "81") Then
                        '仕様有り
                        For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                            If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                               objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "10", "110", "20", "30", "40", "50", "80"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "8" & _
                                                                                   objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Case "11", "111", "21", "31", "41", "51", "81"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & "88" & _
                                                                                   objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End Select
                            End If
                        Next
                    Else
                        '仕様無し
                        If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                Case "10", "110", "20", "30", "40", "50", "80"
                                    If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KB1" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                   strOpPort & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                   strOpPort & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                                Case "11", "111", "21", "31", "41", "51", "81"
                                    If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KB1" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                   strOpPort & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                   strOpPort & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                            End Select
                        Else
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                                If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4KB1" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               Left(strOpPort, 2)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               strOpPort
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If
                    End If

                    '数量セット
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KB1" Then
                            If IsNumeric(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) Then
                                intStationQty = CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                            Else
                                intStationQty = 1
                            End If
                        Else
                            If IsNumeric(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) Then
                                intStationQty = CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                            Else
                                intStationQty = 1
                            End If
                        End If
                    Else
                        intStationQty = 1
                    End If

                    'オプション
                    For intLoopCnt = intIndex To objKtbnStrc.strcSelection.strOpSymbol.Length - 1
                        If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim <> "" Then
                            bolSkipFlg = False
                            strOpKtbn = ""

                            'オプション加算キーセット(1)
                            If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" And _
                               objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim = "H6CE" Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "10", "110", "20", "30", "40", "50", "80"
                                        strOpKtbn = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                    Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 2)
                                    Case "11", "111", "21", "31", "41", "51", "81"
                                        strOpKtbn = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & "8" & CdCst.Sign.Hypen & _
                                                    Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 2)
                                End Select
                            Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                    Case "11"
                                        If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                                            strOpKtbn = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & "8" & CdCst.Sign.Hypen & _
                                                        objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                                        Else
                                            strOpKtbn = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                        objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                                        End If
                                    Case "111", "21", "31", "41", "51", "81"
                                        strOpKtbn = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & "8" & CdCst.Sign.Hypen & _
                                                    objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                                    Case Else
                                        strOpKtbn = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                    objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                                End Select
                            End If

                            '電圧加算キーセット
                            Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 2)
                                Case CdCst.PowerSupply.Div1, CdCst.PowerSupply.Div2
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                                        Case CdCst.PowerSupply.Const1, CdCst.PowerSupply.Const2, CdCst.PowerSupply.Const3
                                            bolSkipFlg = True
                                        Case Else
                                            '2010/08/31 ADD RM0808112(異電圧対応) START--->
                                            If KHKataban.fncVoltageIsStandard(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, _
                                                                            strCountryCd, strOfficeCd) Then
                                                '2010/08/31 ADD RM0808112(異電圧対応) <---END

                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                                                    Case CdCst.PowerSupply.Const4, CdCst.PowerSupply.Const5, CdCst.PowerSupply.Const6
                                                        strOpKtbn = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                                                    Case Else
                                                        strOpKtbn = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "OTH"
                                                End Select
                                            Else
                                                bolSkipFlg = True
                                            End If

                                    End Select
                            End Select

                            If bolSkipFlg = False Then
                                'オプション加算キーセット(2)
                                If intLoopCnt = 3 And objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "9" Then
                                    strOpKtbn = "M" & strOpKtbn
                                End If

                                If Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1) = "8" Then
                                    intQuantity1 = 0
                                    intQuantity2 = 0
                                    For intLoopCnt1 = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                                        Select Case Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt1), 2)
                                            Case "S1"
                                                intQuantity = 1
                                            Case "MP"
                                                intQuantity = 0
                                            Case Else
                                                intQuantity = 2
                                        End Select
                                        If Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt1), 2) <> "MP" Then
                                            intQuantity1 = intQuantity1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt1)
                                        End If
                                        intQuantity2 = intQuantity2 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt1) * intQuantity
                                    Next
                                End If

                                Select Case True
                                    Case Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1) = "8"
                                        If Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 1) = "K" Or _
                                           Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 1) = "P" Or _
                                           Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 1) = "G" Or _
                                           Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 3) = "H10" Or _
                                           Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 3) = "H12" Or _
                                           Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 2) = "H6" Or _
                                           Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 2) = "H8" Or _
                                           objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim = "A" Then
                                            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) = "B" And _
                                               Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 1) = "H" Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = strOpKtbn
                                                decOpAmount(UBound(decOpAmount)) = intStationQty
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = strOpKtbn
                                                decOpAmount(UBound(decOpAmount)) = intQuantity1
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = strOpKtbn
                                            decOpAmount(UBound(decOpAmount)) = intQuantity2
                                        End If
                                    Case Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1) = "1"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpKtbn
                                        decOpAmount(UBound(decOpAmount)) = intStationQty
                                    Case Else
                                        If Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 1) = "K" Or _
                                           Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 1) = "P" Or _
                                           Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 1) = "G" Or _
                                           Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 3) = "H10" Or _
                                           Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 3) = "H12" Or _
                                           Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 2) = "H6" Or _
                                           Left(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim, 2) = "H8" Or _
                                           objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim = "A" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = strOpKtbn
                                            decOpAmount(UBound(decOpAmount)) = intStationQty
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = strOpKtbn
                                            decOpAmount(UBound(decOpAmount)) = intStationQty * 2
                                        End If
                                End Select
                            End If
                        End If

                        Select Case True
                            Case (intLoopCnt = 6 And Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" And _
                                                     objKtbnStrc.strcSelection.strSeriesKataban.Trim <> "M4KB1")
                                intLoopCnt = intLoopCnt + 1
                            Case (intLoopCnt = 7 And objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KB1")
                                intLoopCnt = intLoopCnt + 1
                        End Select
                    Next

                    'オプション加算キーセット(3)
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" And _
                       Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1) = "8" Then
                        Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 4)
                            Case "M5CE", "06CE", "H6CE"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "M" & Mid(objKtbnStrc.strcSelection.strFullKataban.Trim, 2, 5) & _
                                                                                 intStationQty.ToString & "CE"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "M" & Mid(objKtbnStrc.strcSelection.strFullKataban.Trim, 2, 5) & _
                                                                                 intStationQty.ToString
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

    Private Sub subElementChk(ByVal objKtbnStrc As KHKtbnStrc, _
                              ByRef intIndex As Integer, _
                              ByRef strOpPort As String)

        Try

            If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                strOpPort = objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                intIndex = 3
            Else
                strOpPort = objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                intIndex = 4
            End If
            Select Case strOpPort
                Case "", "M5", "06", "M5CE", "06CE", "08", "10", "15"
                Case Else
                    intIndex = intIndex - 1

                    '接続口径セット
                    Select Case True
                        Case (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "3KA1"), _
                             (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4KA1"), _
                             (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M3KA1"), _
                             (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KA1")
                            strOpPort = "M5"
                        Case (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4KA2"), _
                             (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KA2")
                            strOpPort = "06"
                        Case (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4KA3"), _
                             (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KA3")
                            strOpPort = "08"
                        Case (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4KA4"), _
                             (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KA4")
                            strOpPort = "10"
                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KB1" And strOpPort = "M5Y"
                            strOpPort = "M5"
                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KB1" And strOpPort = "H6CE"
                            strOpPort = "06CE"
                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KB1" And strOpPort <> "M5Y" And strOpPort <> "H6CE"
                            strOpPort = "06"
                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KB2" And strOpPort = "H6"
                            strOpPort = "06"
                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KB2" And strOpPort = "06Y"
                            strOpPort = "06"
                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KB2" And strOpPort <> "H6" And strOpPort <> "06Y"
                            strOpPort = "08"
                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KB3" And strOpPort = "H8"
                            strOpPort = "08"
                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KB3" And strOpPort = "08Y"
                            strOpPort = "08"
                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KB3" And strOpPort <> "H8" And strOpPort <> "08Y"
                            strOpPort = "10"
                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4KB4"
                            strOpPort = "10"
                    End Select
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
