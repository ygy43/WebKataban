'************************************************************************************
'*  ProgramID  ：KHPriceR9
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2012/09/27   作成者：Y.Tachi
'*                                      更新日：             更新者：
'*
'*  概要       ：ペンシルシリンダ　ＳＣＰ＊３
'*
'************************************************************************************
Module KHPriceR9

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intIndex As Integer
        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            ''ストローク取得
            'intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc.strcSelection.strSeriesKataban, _
            '                                      objKtbnStrc.strcSelection.strKeyKataban, _
            '                                      CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim), _
            '                                      CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim))

            ''基本価格キー
            'Select Case True
            '    Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "D" Or _
            '         Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "K" Or _
            '         Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "M" Or _
            '         Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "O" Or _
            '         Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "T" Or _
            '         Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "V" Or _
            '         Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "Z"
            '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7) & CdCst.Sign.Hypen & _
            '                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
            '                                                   intStroke.ToString
            '        decOpAmount(UBound(decOpAmount)) = 1
            '    Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 1) = ""
            '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
            '                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
            '                                                   intStroke.ToString
            '        decOpAmount(UBound(decOpAmount)) = 1
            '    Case Else
            '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
            '                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
            '                                                   intStroke.ToString
            '        decOpAmount(UBound(decOpAmount)) = 1
            'End Select

            '基本価格キー
            Select Case True
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 10
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-10"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 11 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 15
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-15"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 16 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 20
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-20"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 21 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 30
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-30"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 31 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 45
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-45"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 46 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 60
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-60"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 61 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 75
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-75"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 76 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 90
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-90"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 91 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 105
                    'RM1305005 2013/06/14 修正
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "6"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-100"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-105"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 106 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 120
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-120"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 121 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 135
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-135"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 136 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 150
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-150"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 151 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 165
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-165"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 166 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 180
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-180"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 181 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 195
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-195"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 196 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 210
                    'RM1305005 2013/06/14 修正
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "10"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-200"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-210"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 211 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 225
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-225"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 226 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 240
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-240"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 241 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 255
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-255"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 256 And CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <= 260
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-260"
                    decOpAmount(UBound(decOpAmount)) = 1
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) >= 261
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban, 5) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-270"
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select


            'ﾏｸﾞﾈｯﾄ加算価格キー
            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "L" Or _
               Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 8, 1) = "L" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3-L"
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '支持形式加算価格キー
            If Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2) = "CB" Or _
               Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2) = "FA" Or _
               Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2) = "LB" Or _
               Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2) = "LS" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "L" Or Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 8, 1) = "L" Then
                intIndex = 5
                If objKtbnStrc.strcSelection.strOpSymbol(intIndex).Trim <> "" Then
                    'スイッチ加算価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & objKtbnStrc.strcSelection.strOpSymbol(intIndex).Trim
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(intIndex + 2).Trim)

                    If objKtbnStrc.strcSelection.strOpSymbol(intIndex + 1).Trim <> "" Then
                        'リード線長さ加算価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & objKtbnStrc.strcSelection.strOpSymbol(intIndex + 1).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(intIndex + 2).Trim)
                    End If
                End If
                intIndex = intIndex + 2
            End If

            'オプション・付属品加算価格キー
            intIndex = 0
            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "L" Then
                If Len(objKtbnStrc.strcSelection.strOpSymbol(8)) <> 0 Then
                    intIndex = 8
                End If
            Else
                If Len(objKtbnStrc.strcSelection.strOpSymbol(5)) <> 0 Then
                    intIndex = 5
                End If
            End If

            If intIndex <> 0 Then
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intIndex), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & strOpArray(intLoopCnt).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                    '食品製造工程向け商品
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "F"
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End Select
                Next
            End If

            'オプション・付属品加算価格キー
            intIndex = 0
            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "L" Then
                If objKtbnStrc.strcSelection.strKeyKataban.Trim = "F" Then
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(9)) <> 0 Then
                        intIndex = 9
                    End If
                End If
            End If

            If intIndex <> 0 Then
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intIndex), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & strOpArray(intLoopCnt).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Next
            End If

            '二次電池加算価格キー
            intIndex = 0
            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 7, 1) = "L" Then
                If objKtbnStrc.strcSelection.strKeyKataban.Trim = "4" Then
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(9)) <> 0 Then
                        intIndex = 9
                    End If
                End If
            Else
                If Len(objKtbnStrc.strcSelection.strOpSymbol(6)) <> 0 Then
                    intIndex = 6
                End If
            End If

            If intIndex <> 0 Then
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intIndex), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "SCP*3" & strOpArray(intLoopCnt).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Next
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
