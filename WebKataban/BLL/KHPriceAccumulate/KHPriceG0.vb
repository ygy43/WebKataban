'************************************************************************************
'*  ProgramID  ：KHPriceG0
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/27   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ハイブリロボ　空圧ロボット用エレメント単軸ユニット　ＨＲＬ－１
'*
'************************************************************************************
Module KHPriceG0

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim intHRLStroke As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '中間STまるめ処理
            Select Case True
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 50
                    intHRLStroke = 50
                Case 51 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) And _
                           CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 75
                    intHRLStroke = 75
                Case 76 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) And _
                           CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 100
                    intHRLStroke = 100
                Case 101 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) And _
                            CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 125
                    intHRLStroke = 125
                Case 126 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) And _
                            CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 150
                    intHRLStroke = 150
                Case 151 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) And _
                            CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 200
                    intHRLStroke = 200
                Case 201 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) And _
                            CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 250
                    intHRLStroke = 250
                Case 251 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) And _
                            CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 300
                    intHRLStroke = 300
                Case 301 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) And _
                            CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 350
                    intHRLStroke = 350
                Case 351 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) And _
                            CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 400
                    intHRLStroke = 400
                Case 401 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) And _
                            CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 450
                    intHRLStroke = 450
                Case 451 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) And _
                            CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 500
                    intHRLStroke = 500
                Case 501 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) And _
                            CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <= 550
                    intHRLStroke = 550
                Case 551 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                    intHRLStroke = 600
            End Select

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = "HRL-1-" & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & intHRLStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1

            'オプション加算価格キー(落下防止機構)
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "HRL-1-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                                                                      objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'オプション加算価格キー(フランジ)
            '基本形状判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "F", "LF"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "HRL-1-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                          objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション加算価格キー(スイッチ)
            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "HRL-1-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)

                'オプション加算価格キー(レール)
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "HRL-1-RAIL"
                decOpAmount(UBound(decOpAmount)) = 1

                'オプション加算価格キー(リード線長さ)
                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                    '2010/10/19 RM1010017(11月VerUP:HRLシリーズ) START--->
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        Case "T2YD"
                            strOpRefKataban(UBound(strOpRefKataban)) = "HRL-T2YD-" & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        Case "T2YDT"
                            strOpRefKataban(UBound(strOpRefKataban)) = "HRL-T2YDT-" & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        Case Else
                            strOpRefKataban(UBound(strOpRefKataban)) = "HRL-1-" & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    End Select
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)

                    'ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    'strOpRefKataban(UBound(strOpRefKataban)) = "HRL-1-" & objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    'decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                    '2010/10/19 RM1010017(11月VerUP:HRLシリーズ) <---END
                End If
            End If

            '2011/10/24 ADD RM1110032(11月VerUP:二次電池) START--->
            '二次電池用
            If objKtbnStrc.strcSelection.strKeyKataban = "2" Then
                'スイッチ加算価格キー
                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & CdCst.Sign.Hypen & _
                                                            "SW" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)

                End If

                '二次電池加算価格キー
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban & CdCst.Sign.Hypen & _
                                                        objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If
            '2011/10/24 ADD RM1110032(11月VerUP:二次電池) <---END

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
