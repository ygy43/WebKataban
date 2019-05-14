'************************************************************************************
'*  ProgramID  ：KHPrice34
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/07   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：セレックスバルブ　４Ｆ／Ｍ４Ｆ（防爆）
'*
'************************************************************************************
Module KHPrice34

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strCountryCd As String = Nothing, _
                                   Optional ByRef strOfficeCd As String = Nothing)



        Dim strStdVoltageFlag As String = CdCst.VoltageDiv.Standard
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim strWkKataban As String = ""
        Dim intQuantity1 As String = 0
        Dim intQuantity2 As String = 0

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            'RM1312084 2013/12/25 追加
            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Case "EX"
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        strWkKataban = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 3) & "EX"
                    Else
                        strWkKataban = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "EX"
                    End If

                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                            'サブプレート加算価格キー(M4F*のみ)
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & "SP" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            '仕様入力キー(M4F*のみ)
                            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                                If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                                    If Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 3) = "4F3" And _
                                       Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 2) <> "MP" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt) & "M"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    End If

                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 1)
                                        Case "1"
                                            intQuantity1 = intQuantity1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Case "2", "3", "4", "5"
                                            intQuantity1 = intQuantity1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                    End Select

                                End If
                            Next
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 4) & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 1)
                            '    Case "1"
                            '        intQuantity1 = intQuantity1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            '    Case "2", "3", "4", "5"
                            '        intQuantity1 = intQuantity1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                            'End Select
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                Case "1"
                                    intQuantity1 = intQuantity1 + objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                Case "2", "3", "4", "5"
                                    intQuantity1 = intQuantity1 + objKtbnStrc.strcSelection.strOpSymbol(8).Trim * 2
                            End Select
                        End If
                    Else

                        'NAMUR規格対応品対応  2016/11/21 変更 松原
                        'シリーズ型番が「4F1」「4F3」の場合のみ以下の処理を実施
                        If objKtbnStrc.strcSelection.strSeriesKataban = "4F1" Or objKtbnStrc.strcSelection.strSeriesKataban = "4F3" Then
                            If objKtbnStrc.strcSelection.strKeyKataban = "A" Or objKtbnStrc.strcSelection.strKeyKataban = "N" Or objKtbnStrc.strcSelection.strKeyKataban = "Z" Then
                                'NAMUR規格対応品のみ、単価キーの末尾に「-NM」をつける
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "EX" & "-NM"
                            Else
                                'それ以外は通常通り処理
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "EX"
                            End If
                        Else
                            'それ以外は通常通り処理
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "EX"
                        End If

                        ''基本価格
                        'ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                        '                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                        '                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "EX"
                        'Select Case True
                        '    Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "9" And _
                        '         Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "6"
                        '        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "D00"
                        '    Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "9" And _
                        '         Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "7"
                        '        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "E00"
                        'End Select
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    Select Case True
                        Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M"
                            intQuantity2 = intQuantity1
                        Case Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1) = "1"
                            intQuantity2 = 1
                        Case Else
                            intQuantity2 = 2
                    End Select

                    '電圧加算価格キー
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        '2016/08/23 ADD RM1608022 START--->
                        If objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                Case "4", "5", "6", "C", "E", "G", "H", "L"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & "ACOPT"
                                    decOpAmount(UBound(decOpAmount)) = intQuantity2
                                Case "3", "7", "8", "9", "A", "B"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & "DCOPT"
                                    decOpAmount(UBound(decOpAmount)) = intQuantity2
                            End Select
                        Else '2016/08/23 ADD RM1608022 <---End
                            strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(10).Trim, _
                                                                           strCountryCd, strOfficeCd)
                            Select Case strStdVoltageFlag
                                Case CdCst.VoltageDiv.Standard
                                Case CdCst.VoltageDiv.Options
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2) & "OPT"
                                    decOpAmount(UBound(decOpAmount)) = intQuantity2
                                Case CdCst.VoltageDiv.Other
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & Left(objKtbnStrc.strcSelection.strOpSymbol(10).Trim, 2) & "OTH"
                                    decOpAmount(UBound(decOpAmount)) = intQuantity2
                            End Select
                        End If
                    Else
                        strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                                                                       strCountryCd, strOfficeCd)
                        Select Case strStdVoltageFlag
                            Case CdCst.VoltageDiv.Standard
                            Case CdCst.VoltageDiv.Options
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & Left(objKtbnStrc.strcSelection.strOpSymbol(8).Trim, 2) & "OPT"
                                decOpAmount(UBound(decOpAmount)) = intQuantity2
                            Case CdCst.VoltageDiv.Other
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & Left(objKtbnStrc.strcSelection.strOpSymbol(8).Trim, 2) & "OTH"
                                decOpAmount(UBound(decOpAmount)) = intQuantity2
                        End Select
                    End If
                    '2016/08/23 ADD RM1608022 <---End 

                    'オプション加算価格キー

                    '2016/11/21 追加 松原
                    'オプション加算前にNAMUR規格対応品かチェックを行い、そうだった場合変数strWkKatabanに「-NM-」を付与する
                    'シリーズ型番が「4F1」「4F3」の場合のみ以下の処理を実施
                    If objKtbnStrc.strcSelection.strSeriesKataban = "4F1" Or objKtbnStrc.strcSelection.strSeriesKataban = "4F3" Then
                        If objKtbnStrc.strcSelection.strKeyKataban = "A" Or objKtbnStrc.strcSelection.strKeyKataban = "N" Or objKtbnStrc.strcSelection.strKeyKataban = "Z" Then
                            strWkKataban = strWkKataban & "-NM-"
                        End If
                    End If

                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                If strOpArray(intLoopCnt).Trim = "H" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & strOpArray(intLoopCnt).Trim
                                End If

                                If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case "N", "NC", "NO"
                                            decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                                        Case Else
                                            decOpAmount(UBound(decOpAmount)) = intQuantity2
                                    End Select
                                Else
                                    If strOpArray(intLoopCnt).Trim = "R" Then
                                        decOpAmount(UBound(decOpAmount)) = intQuantity2
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                                End If
                        End Select
                    Next

                    '排気取付加算価格キー(M4F*のみ)
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 3) & "EX" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '承認機関加算価格キー '2016/08/23 ADD RM1608022 START
                    'NAMUR規格対応品対応による追加　2016/11/21 追加 松原 
                    If objKtbnStrc.strcSelection.strKeyKataban = "Y" Or objKtbnStrc.strcSelection.strKeyKataban = "Z" Then
                        'If objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                        If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                            If objKtbnStrc.strcSelection.strOpSymbol(11).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "4FEX-" & objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                                'decOpAmount(UBound(decOpAmount)) = 1
                                decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) '2016/09/09 RM1609041 連数を積み上げ
                            End If
                        Else
                            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "4FEX-" & objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If
                    End If '2016/08/23 ADD RM1608022 <---End 


                Case Else
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        strWkKataban = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 3) & "E"
                    Else
                        strWkKataban = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "E"
                    End If

                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        'サブプレート加算価格キー(M4F*のみ)
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & "SP" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                            '仕様入力キー(M4F*のみ)
                            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                                If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 1)
                                        Case "1"
                                            intQuantity1 = intQuantity1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Case "2", "3", "4", "5"
                                            intQuantity1 = intQuantity1 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                    End Select
                                End If
                            Next
                        Else
                            Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1)
                                Case "3"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 3) & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "0" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                Case "4", "5"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 3) & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "9" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                Case "6"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 3) & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "9" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 3) & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "9" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                            End Select

                            Select Case Mid(strOpRefKataban(UBound(strOpRefKataban)), 4, 1)
                                Case "1"
                                    intQuantity1 = CDec(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                Case Else
                                    intQuantity1 = CDec(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) * 2
                            End Select
                        End If
                    Else
                        '基本価格

                        'NAMUR規格対応品対応  2016/11/21 変更 松原
                        'シリーズ型番が「4F1」「4F3」の場合のみ以下の処理を実施
                        If objKtbnStrc.strcSelection.strSeriesKataban = "4F1" Or objKtbnStrc.strcSelection.strSeriesKataban = "4F3" Then

                            If objKtbnStrc.strcSelection.strKeyKataban = "A" Or objKtbnStrc.strcSelection.strKeyKataban = "N" Or objKtbnStrc.strcSelection.strKeyKataban = "Z" Then
                                'NAMUR規格対応品の場合は「-NM」をつける
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                          objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                          objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "E" & "-NM"
                            Else
                                '通常品の場合処理はそのまま
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                          objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                          objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "E"
                            End If

                        Else

                            '通常品の場合処理はそのまま
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                      objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                      objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "E"

                        End If

                        'ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        'strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                        '                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                        '                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "E"

                        Select Case True
                            Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "9" And _
                                 Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "6"
                                '↓RM1401080 2014/01/23
                                'strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "D00"
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1) = "9" And _
                                 Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "7"
                                '↓RM1401080 2014/01/23
                                'strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "E00"
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        End Select
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    Select Case True
                        Case Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M"
                            intQuantity2 = intQuantity1
                        Case Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1) = "1"
                            intQuantity2 = 1
                        Case Else
                            intQuantity2 = 2
                    End Select

                    '外部導線引込式加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                        'NAMUR規格対応品対応  2016/11/21 変更 松原
                        'シリーズ型番が「4F1」「4F3」の場合のみ以下の処理を実施
                        If objKtbnStrc.strcSelection.strSeriesKataban = "4F1" Or objKtbnStrc.strcSelection.strSeriesKataban = "4F3" Then

                            If objKtbnStrc.strcSelection.strKeyKataban = "A" Or objKtbnStrc.strcSelection.strKeyKataban = "N" Or objKtbnStrc.strcSelection.strKeyKataban = "Z" Then
                                '「-NM-」を付与する
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & "-NM-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                decOpAmount(UBound(decOpAmount)) = intQuantity2
                            Else
                                '通常通り処理を行う
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                decOpAmount(UBound(decOpAmount)) = intQuantity2

                            End If
                        Else
                            '通常通り処理を行う
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            decOpAmount(UBound(decOpAmount)) = intQuantity2

                        End If

                        'ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        'strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        'decOpAmount(UBound(decOpAmount)) = intQuantity2
                    End If

                    '絶縁種別加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then

                        'NAMUR規格対応品対応  2016/11/21 変更 松原
                        'シリーズ型番が「4F1」「4F3」の場合のみ以下の処理を実施
                        If objKtbnStrc.strcSelection.strSeriesKataban = "4F1" Or objKtbnStrc.strcSelection.strSeriesKataban = "4F3" Then

                            If objKtbnStrc.strcSelection.strKeyKataban = "A" Or objKtbnStrc.strcSelection.strKeyKataban = "N" Or objKtbnStrc.strcSelection.strKeyKataban = "Z" Then
                                '「-NM-」を付与する
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & "-NM-" & objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = intQuantity2
                            Else
                                '通常通り処理を行う
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = intQuantity2
                            End If

                        Else
                            '通常通り処理を行う
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                            decOpAmount(UBound(decOpAmount)) = intQuantity2

                        End If

                        'ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        'strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                        'decOpAmount(UBound(decOpAmount)) = intQuantity2
                    End If

                    '電圧加算価格キー
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        '2010/08/31 ADD RM0808112(異電圧対応) START--->
                        strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(11).Trim, _
                                                                       strCountryCd, strOfficeCd)
                        'strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                        '                                               objKtbnStrc.strcSelection.strOpSymbol(11).Trim)
                        '2010/08/31 ADD RM0808112(異電圧対応) <---END
                        Select Case strStdVoltageFlag
                            Case CdCst.VoltageDiv.Standard
                            Case CdCst.VoltageDiv.Options
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & Left(objKtbnStrc.strcSelection.strOpSymbol(11).Trim, 2) & "OPT"
                                decOpAmount(UBound(decOpAmount)) = intQuantity2
                            Case CdCst.VoltageDiv.Other
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & Left(objKtbnStrc.strcSelection.strOpSymbol(11).Trim, 2) & "OTH"
                                decOpAmount(UBound(decOpAmount)) = intQuantity2
                        End Select
                    Else
                        '2010/08/31 ADD RM0808112(異電圧対応) START--->
                        strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                                                       strCountryCd, strOfficeCd)
                        'strStdVoltageFlag = KHKataban.fncVoltageInfoGet(objKtbnStrc, _
                        '                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                        '2010/08/31 ADD RM0808112(異電圧対応) <---END
                        Select Case strStdVoltageFlag
                            Case CdCst.VoltageDiv.Standard
                            Case CdCst.VoltageDiv.Options
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & Left(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, 2) & "OPT"
                                decOpAmount(UBound(decOpAmount)) = intQuantity2
                            Case CdCst.VoltageDiv.Other
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & Left(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, 2) & "OTH"
                                decOpAmount(UBound(decOpAmount)) = intQuantity2
                        End Select
                    End If

                    'オプション加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)

                    '2016/11/21 追加 松原
                    'オプション加算前にNAMUR規格対応品かチェックを行い、そうだった場合変数strWkKatabanに「-NM-」を付与する
                    'シリーズ型番が「4F1」「4F3」の場合のみ以下の処理を実施
                    If objKtbnStrc.strcSelection.strSeriesKataban = "4F1" Or objKtbnStrc.strcSelection.strSeriesKataban = "4F3" Then
                        If objKtbnStrc.strcSelection.strKeyKataban = "A" Or objKtbnStrc.strcSelection.strKeyKataban = "N" Or objKtbnStrc.strcSelection.strKeyKataban = "Z" Then
                            strWkKataban = strWkKataban & "-NM-"
                        End If
                    End If

                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                If strOpArray(intLoopCnt).Trim = "H" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strWkKataban & strOpArray(intLoopCnt).Trim
                                End If

                                If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case "N", "NC", "NO"
                                            decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                        Case Else
                                            decOpAmount(UBound(decOpAmount)) = intQuantity2
                                    End Select
                                Else
                                    If strOpArray(intLoopCnt).Trim = "R" Then
                                        decOpAmount(UBound(decOpAmount)) = intQuantity2
                                    Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                                End If
                        End Select
                    Next

                    '排気取付加算価格キー(M4F*のみ)
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "M" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 3) & "E" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select
        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
