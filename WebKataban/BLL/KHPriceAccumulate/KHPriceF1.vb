'************************************************************************************
'*  ProgramID  ：KHPriceF1
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/26   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ニューハンドリングシステム　ＮＨＳ－Ｈ
'*
'************************************************************************************
Module KHPriceF1

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim intStroke As Integer

        Dim intZStroke As Integer
        Dim strXSymbol As String
        Dim strKSymbol As String
        
        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '可搬質量を変換する(1)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "1005", "1007"
                    strXSymbol = "10"
                Case "1505", "1507", "1510", "1512"
                    strXSymbol = "15"
                Case "3010", "3012", "3020"
                    strXSymbol = "30"
                Case "5010", "5012", "5020", "5033"
                    strXSymbol = "50"
                Case Else
                    strXSymbol = "10"
            End Select

            '可搬質量を変換する(2)
            Select Case Mid(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 3, 2)
                Case "05"
                    strKSymbol = "05"
                Case "07"
                    strKSymbol = "10"
                Case "10"
                    strKSymbol = "15"
                Case "12"
                    strKSymbol = "15H"
                Case "20"
                    strKSymbol = "25"
                Case "33"
                    strKSymbol = "50"
                Case Else
                    strKSymbol = ""
            End Select

            'Z軸ストロークの値を求める(中間STまるめ処理)
            Select Case True
                Case CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 50
                    intZStroke = 50
                Case 51 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) And CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 75
                    intZStroke = 75
                Case 76 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) And CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 100
                    intZStroke = 100
                Case 101 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) And CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 125
                    intZStroke = 125
                Case 126 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) And CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 150
                    intZStroke = 150
                Case 151 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) And CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 200
                    intZStroke = 200
                Case 201 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) And CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 250
                    intZStroke = 250
                Case 251 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) And CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 300
                    intZStroke = 300
                Case 301 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) And CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 350
                    intZStroke = 350
                Case 351 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) And CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 400
                    intZStroke = 400
                Case 401 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) And CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 450
                    intZStroke = 450
                Case 451 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) And CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 500
                    intZStroke = 500
                Case 501 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) And CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <= 550
                    intZStroke = 550
                Case 551 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim)
                    intZStroke = 600
            End Select

            'X軸処理
            '(X軸)基本価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "1"
                    'ストローク設定
                    intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                          CInt(objKtbnStrc.strcSelection.strOpSymbol(1).Trim), _
                                                          CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim))

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR" & CdCst.Sign.Hypen & _
                                                               strXSymbol & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "2"
                    'ストローク設定
                    intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                          CInt(objKtbnStrc.strcSelection.strOpSymbol(1).Trim), _
                                                          CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim))

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR" & CdCst.Sign.Hypen & _
                                                               strXSymbol & CdCst.Sign.Hypen & _
                                                               intStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション加算価格キー(2ヘッド)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "2"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR" & CdCst.Sign.Hypen & "2HEAD" & CdCst.Sign.Hypen & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '(X軸)オプション加算価格キー(ストローク調整ブロック)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                Case "L", "R", "D"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR" & CdCst.Sign.Hypen & "STAB" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                               strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '(X軸)オプション加算価格キー(スピードコントローラ)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                Case "3", "4"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR" & CdCst.Sign.Hypen & "SCLB" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & _
                                                               strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '(X軸)オプション加算価格キー(ケーブルベア)
            'ケーブルベア判定
            Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                Case "B", "T"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR-CABLE-B-" & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "W", "Y"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR-CABLE-W-" & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '(X軸)オプション加算価格キー(スイッチ)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                Case "A"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'Z軸処理
            '(Z軸)基本価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "1"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NHS-H-" & strKSymbol & CdCst.Sign.Hypen & intZStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "2"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NHS-H-" & strKSymbol & CdCst.Sign.Hypen & intZStroke.ToString
                    decOpAmount(UBound(decOpAmount)) = 2
            End Select

            '(Z軸)オプション加算価格キー(スピードコントローラ)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                Case "3", "4"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "1"    '１ヘッド
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-H-SC-" & strKSymbol
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2"    '２ヘッド
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-H-SC-" & strKSymbol
                            decOpAmount(UBound(decOpAmount)) = 2
                    End Select
            End Select

            '(Z軸)オプション加算価格キー(ケーブルベア)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                Case "T", "Y"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-H-CABLE-" & objKtbnStrc.strcSelection.strOpSymbol(9).Trim & CdCst.Sign.Hypen & strKSymbol
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-H-CABLE-" & objKtbnStrc.strcSelection.strOpSymbol(9).Trim & CdCst.Sign.Hypen & strKSymbol
                            decOpAmount(UBound(decOpAmount)) = 2
                    End Select
            End Select

            '(Z軸)オプション加算価格キー(スイッチ)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                Case "A"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-H-" & objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-H-" & objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            decOpAmount(UBound(decOpAmount)) = 2
                    End Select
            End Select

            '(Z軸)オプション加算価格キー(落下防止機構)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                Case "Q"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-H-" & objKtbnStrc.strcSelection.strOpSymbol(11).Trim & CdCst.Sign.Hypen & strKSymbol
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-H-" & objKtbnStrc.strcSelection.strOpSymbol(11).Trim & CdCst.Sign.Hypen & strKSymbol
                            decOpAmount(UBound(decOpAmount)) = 2
                    End Select
            End Select

            '(Z軸)オプション加算価格キー(アタッチメント)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "1"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NHS-H-ATATCHMENT-" & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "2"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NHS-H-ATATCHMENT-" & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 2
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
