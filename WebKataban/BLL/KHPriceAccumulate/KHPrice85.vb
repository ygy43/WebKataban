'************************************************************************************
'*  ProgramID  ：KHPrice23
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/21   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ペンシルシリンダ複動形　ＳＣＰＤ２／ＳＣＰＤ２－Ｌ
'*
'************************************************************************************
Module KHPrice85

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer

        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False
        Dim bolOptionP4 As Boolean = False      'RM1001045 2010/02/23 Y.Miura　二次電池対応

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'RM1001045 2010/02/23 Y.Miura 二次電池機器追加
            If objKtbnStrc.strcSelection.strOpSymbol.Length > 9 Then
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(9), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case "P4", "P40"
                            bolOptionP4 = True
                    End Select
                Next
            End If


            'ストローク取得
            intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim))

            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "SCPG2", "SCPG2-L", "SCPG2-X", "SCPG2-XL", "SCPG2-Y", "SCPG2-YL"
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "SCPG2-X", "SCPG2-XL"
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & "-X" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                       intStroke.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "SCPG2-Y", "SCPG2-YL"
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & "-Y" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                       intStroke.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                       intStroke.ToString
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'シリーズオプション加算価格キー(2)
                    If Mid(objKtbnStrc.strcSelection.strSeriesKataban, 7, 1) = "L" Or _
                       Mid(objKtbnStrc.strcSelection.strSeriesKataban, 8, 1) = "L" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2-L"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '支持形式加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'スイッチ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)

                        'リード線長さ加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                        End If

                    End If

                    'オプション・付属品加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(9), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCPG2" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                Case Else
                    'バリエーション(微速)加算価格キー
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "F"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "6"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 30
                                            'ストローク10～30
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR10" & CdCst.Sign.Hypen & "30"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 31 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 60
                                            'ストローク31～60
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR31" & CdCst.Sign.Hypen & "60"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 61
                                            'ストローク61～100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR61" & CdCst.Sign.Hypen & "100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Case "10"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 45
                                            'ストローク10～45
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR10" & CdCst.Sign.Hypen & "45"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 46 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 100
                                            'ストローク46～100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR46" & CdCst.Sign.Hypen & "100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 101
                                            'ストローク101～200
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR101" & CdCst.Sign.Hypen & "200"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Case "16"
                                    Select Case True
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 45
                                            'ストローク10～45
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR10" & CdCst.Sign.Hypen & "45"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 46 And _
                                             CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 100
                                            'ストローク46～100
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR46" & CdCst.Sign.Hypen & "100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) >= 101
                                            'ストローク101～260
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "STR101" & CdCst.Sign.Hypen & "260"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                            End Select
                    End Select

                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1

                    'シリーズオプション加算価格キー(2)
                    If Mid(objKtbnStrc.strcSelection.strSeriesKataban, 7, 1) = "L" Or _
                       Mid(objKtbnStrc.strcSelection.strSeriesKataban, 8, 1) = "L" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2-L"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '支持形式加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "CB" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'スイッチ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)

                        'リード線長さ加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                        End If

                        'RM1001045 2010/02/23 Y.Miura 二次電池機器追加
                        'P4加算
                        If bolOptionP4 Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & "-SW-P4"
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                        End If
                    End If

                    'オプション・付属品加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(9), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*2" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next

                    'クリーン仕様加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(11).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(11).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select
        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
