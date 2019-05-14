'************************************************************************************
'*  ProgramID  ：KHPrice45
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/27   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：４Ｌ／Ｍ４Ｌ／４Ｓ＊０／Ｍ４Ｓ＊０／３Ｍ＊０／Ｍ３Ｍ＊０／３Ｐ／Ｍ３Ｐ
'*
'*  更新履歴   ：                       更新日：2009/03/06   更新者：T.Yagyu
'*               ・RM0902053: M4S 配管接続オプション加算の数量 バルブ数→連数
'************************************************************************************
Module KHPrice45

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strStdVoltageFlag As String = CdCst.VoltageDiv.Standard
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim strSrzKbn As String = ""
        Dim intIndex As Integer = 0
        Dim intQuantity1 As Integer = 0
        Dim intQuantity2 As Integer = 0
        Dim intStationQty As Integer = 0
        Dim strKeyKat As String = ""

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                strSrzKbn = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1)
            Else
                strSrzKbn = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 1)
            End If

            If objKtbnStrc.strcSelection.strSpecNo.Trim <> "" Then
                '仕様有り
                'サブプレート
                If strSrzKbn = "L" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4)
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5)
                End If
                If strSrzKbn = "M" Then
                    intIndex = 6
                Else
                    intIndex = 7
                End If
                If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4L3" Then
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "T" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "T1" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-T"
                    Else
                        If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-DIN"
                        End If
                    End If
                End If
                If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4LB2" Then
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-" & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & "-SP" & objKtbnStrc.strcSelection.strOpSymbol(intIndex).Trim

                Else
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-SP" & objKtbnStrc.strcSelection.strOpSymbol(intIndex).Trim
                End If

                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                decOpAmount(UBound(decOpAmount)) = 1

                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                    For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                        If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                            If strSrzKbn = "L" Then
                                If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4LB2" Then
                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 5, 1) = "1" Then
                                        intQuantity1 = intQuantity1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        intQuantity2 = intQuantity2 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 6, 1) <> "M" Then
                                            intQuantity1 = intQuantity1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                            intQuantity2 = intQuantity2 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                Else
                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 1) = "1" Then
                                        intQuantity1 = intQuantity1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        intQuantity2 = intQuantity2 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 6, 1) <> "M" Then
                                            intQuantity1 = intQuantity1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                            intQuantity2 = intQuantity2 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                                End If
                            Else
                                If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 7, 1) <> "M" Then
                                    If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 5, 1) = "1" Then
                                        intQuantity1 = intQuantity1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        intQuantity1 = intQuantity1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                    End If
                                    intQuantity2 = intQuantity2 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                            End If
                        End If
                    Next
                Else
                    Select Case objKtbnStrc.strcSelection.strSpecNo.Trim
                        Case "89", "90"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 3) & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "9"
                        Case "98"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "9"
                        Case Else
                            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M3MA0" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "9-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 4) & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "9"
                            End If
                    End Select
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                        Case "M3MA0", "M3MB0"
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                        Case Else
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                    End Select
                    If strSrzKbn = "L" Then
                        If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4LB2" Then
                            '201501月次更新
                            If Mid(strOpRefKataban(UBound(strOpRefKataban)), 5, 1) = "1" Then
                                intQuantity1 = decOpAmount(UBound(decOpAmount))
                                intQuantity2 = decOpAmount(UBound(decOpAmount))
                            Else
                                'If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 6, 1) <> "M" Then
                                intQuantity1 = decOpAmount(UBound(decOpAmount)) * 2
                                intQuantity2 = decOpAmount(UBound(decOpAmount))
                                'End If
                            End If
                        ElseIf Mid(strOpRefKataban(UBound(strOpRefKataban)), 4, 1) = "1" Then
                            intQuantity1 = decOpAmount(UBound(decOpAmount))
                            intQuantity2 = decOpAmount(UBound(decOpAmount))
                        Else
                            'If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 6, 1) <> "M" Then
                            intQuantity1 = decOpAmount(UBound(decOpAmount)) * 2
                            intQuantity2 = decOpAmount(UBound(decOpAmount))
                            'End If
                        End If
                    Else
                        If Mid(strOpRefKataban(UBound(strOpRefKataban)), 7, 1) <> "M" Then
                            If Mid(strOpRefKataban(UBound(strOpRefKataban)), 5, 1) = "1" Then
                                intQuantity1 = decOpAmount(UBound(decOpAmount))
                            Else
                                intQuantity1 = decOpAmount(UBound(decOpAmount)) * 2
                            End If
                            intQuantity2 = decOpAmount(UBound(decOpAmount))
                        End If
                    End If
                End If

                If strSrzKbn = "M" Then
                    intIndex = 6
                Else
                    intIndex = 7
                End If
                intStationQty = objKtbnStrc.strcSelection.strOpSymbol(intIndex).Trim
                strKeyKat = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2)
            Else
                '仕様無し
                '基本価格キー
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
                strKeyKat = objKtbnStrc.strcSelection.strSeriesKataban.Trim

                If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "3MA0" Then
                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                End If
                If strSrzKbn = "L" Or objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Then
                    intQuantity1 = 1
                Else
                    intQuantity1 = 2
                End If
                intStationQty = 1
            End If

            If strSrzKbn = "L" Or _
               strSrzKbn = "M" And Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) <> "M" Or _
               strSrzKbn = "M" And Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" And (objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "M3" Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "M5") Or _
               objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M3MA0" And objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "T4" Or _
               strSrzKbn = "S" And (objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "00" Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "M3" Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "M5") Or _
               strSrzKbn = "P" And (objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "00" Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "M5" Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "06" Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "08" Or _
               objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "06Y" Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "06A" Or objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "06B") Then
            Else
                '加算価格
                '配管接続加算
                If (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4SA0" Or _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim = "3PA1" Or _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim = "3PA2" Or _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim = "3PB1") And _
                      objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "9" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & strKeyKat
                    If strSrzKbn = "L" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), 4)
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), 5)
                    End If
                Else
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5)
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4)
                    End If
                    If strSrzKbn = "L" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4)
                    End If
                End If
                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim

                '2009/03/06 T.Y 変更前
                'If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4SA0" And objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "M3" Or _
                '  (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M3PA1" Or _
                '   objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M3PA2") And objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "06" Then
                '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                '    decOpAmount(UBound(decOpAmount)) = intQuantity2
                'ElseIf objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4SB0" Then
                '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                '    decOpAmount(UBound(decOpAmount)) = intQuantity2
                'Else
                '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                '    decOpAmount(UBound(decOpAmount)) = intStationQty
                'End If

                '2009/03/06 T.Y 変更後
                If (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M3PA1" Or _
                   objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M3PA2") And objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "06" Then
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = intQuantity2
                ElseIf objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4SA0" Or objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4SB0" Then
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = intStationQty
                Else
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = intStationQty
                End If
            End If

            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "" Or strSrzKbn = "L" And objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "M6" Then
            Else
                '手動装置加算価格
                If strSrzKbn = "L" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(strKeyKat, 3)
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(strKeyKat, 4)
                End If
                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim

                If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4L3" Then
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "6" Then
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 2
                    End If
                Else
                    If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4L2" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Then
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                    ElseIf objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4LB2" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Then
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                    Else
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = intQuantity1
                    End If
                End If
            End If

            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "" Or _
              (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4L3" Or _
               objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4L3") Then
            Else
                If strSrzKbn = "S" And Mid(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 3, 1) = "T" Then
                    '省配線の配線ブロック
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(strKeyKat, 4) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
                '電線接続加算
                Select Case True
                    Case strSrzKbn = "L"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strKeyKat, 3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    Case strSrzKbn = "S" And Mid(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 3, 1) = "T"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strKeyKat, 4) & CdCst.Sign.Hypen & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 2)
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strKeyKat, 4) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                End Select
                If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4L2" Then
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = intQuantity2
                Else
                    If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4L2" Then
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    ElseIf objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4LB2" Then
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = intQuantity1
                    End If
                End If
            End If

            If strSrzKbn <> "M" And Left(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 1) <> "" Then
                'オプション加算
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else
                            If strSrzKbn = "L" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(strKeyKat, 3) & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(strKeyKat, 4) & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                            End If
                            If strSrzKbn = "L" Then
                                If strOpArray(intLoopCnt).Trim = "N" Then
                                    If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4L2" Then
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4L3" Then
                                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Or _
                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "6" Then
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Else
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                decOpAmount(UBound(decOpAmount)) = 2
                                            End If
                                        Else
                                            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4L2" Then
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                decOpAmount(UBound(decOpAmount)) = intStationQty
                                            Else
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                decOpAmount(UBound(decOpAmount)) = intQuantity1
                                            End If
                                        End If
                                    End If
                                Else
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    decOpAmount(UBound(decOpAmount)) = intStationQty
                                End If
                            Else
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                decOpAmount(UBound(decOpAmount)) = intQuantity1
                            End If
                    End Select
                Next
            End If

            '電圧加算
            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOpElementDiv.Length - 1
                If objKtbnStrc.strcSelection.strOpElementDiv(intLoopCnt) = CdCst.ElementDiv.Voltage Then
                    strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim)
                    Select Case strStdVoltageFlag
                        Case CdCst.VoltageDiv.Standard
                        Case Else
                            If strSrzKbn = "L" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(strKeyKat, 3)
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(strKeyKat, 4)
                            End If
                            If strStdVoltageFlag = CdCst.VoltageDiv.Options Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-OPT"
                            Else
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-OTH"
                            End If
                            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4L3" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Or _
                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "6" Then
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    decOpAmount(UBound(decOpAmount)) = 2
                                End If
                            Else
                                If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4L2" Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "1" Then
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    End If
                                Else
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    decOpAmount(UBound(decOpAmount)) = intQuantity1
                                End If
                            End If
                    End Select
                End If
            Next

            If strSrzKbn <> "L" And strSrzKbn <> "M" Or _
               Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" And strSrzKbn <> "M" Or _
               Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) <> "M" Or _
               objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "M3" And objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "M5" And objKtbnStrc.strcSelection.strSeriesKataban.Trim <> "M3MA0" Or _
               objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "T4" And strSrzKbn <> "S" Or _
               objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "00" And objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "M3" And objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "M5" And strSrzKbn <> "P" Or _
               objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "00" And objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "M5" And objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "06" And objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "08" And _
               objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "06Y" And objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "06A" And objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "06B" Then
            Else
                '配管接続加算
                If (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "4SA0" Or _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim = "3PA1" Or _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim = "3PA2" Or _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim = "3PB1") And _
                    objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "9" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "M" & strKeyKat
                    If strSrzKbn = "L" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), 4)
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)), 5)
                    End If
                Else
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5)
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4)
                    End If
                    If strSrzKbn = "L" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4)
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3)
                    End If
                End If
                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim

                If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M4SA0" And objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "M3" Or _
                  (objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M3PA1" Or objKtbnStrc.strcSelection.strSeriesKataban.Trim = "M3PA2") And _
                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "06" Then
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = intQuantity2
                Else
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = intStationQty
                End If
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
