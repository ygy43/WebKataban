'************************************************************************************
'*  ProgramID  ：KHPriceF1
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/26   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ニューハンドリングシステム　ＮＨＳ－Ｃ
'*
'************************************************************************************
Module KHPriceF2

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim intStroke As Integer

        Dim strXSymbol As String
        Dim strKSymbol As String

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '可搬質量を変換する(1)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "1004", "1006"
                    strXSymbol = "10"
                Case "1504", "1506", "1510", "1512"
                    strXSymbol = "15"
                Case "3010", "3012"
                    strXSymbol = "30"
                Case "5010", "5012"
                    strXSymbol = "50"
                Case Else
                    strXSymbol = ""
            End Select

            '可搬質量を変換する(2)
            Select Case Mid(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 3, 2)
                Case "04"
                    strKSymbol = "12"
                Case "06"
                    strKSymbol = "16"
                Case "10"
                    strKSymbol = "20"
                Case "12"
                    strKSymbol = "25"
                Case Else
                    strKSymbol = ""
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
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR" & "-2HEAD-" & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '(X軸)オプション加算価格キー(ストローク調整ブロック)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                Case "L", "R", "D"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR-STAB-" & objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '(X軸)オプション加算価格キー(スピードコントローラ)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                Case "3", "4"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR-SCLB-" & objKtbnStrc.strcSelection.strOpSymbol(8).Trim & CdCst.Sign.Hypen & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '(X軸)オプション加算価格キー(ケーブルベア)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                Case "B", "W"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR-CABLE-" & objKtbnStrc.strcSelection.strOpSymbol(9).Trim & CdCst.Sign.Hypen & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '(X軸)オプション加算価格キー(スイッチ)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                Case "A"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NSR" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'Z軸処理
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "1"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NHS-C-" & strKSymbol & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "2"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NHS-C-" & strKSymbol & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 2
            End Select

            '(Z軸)オプション加算価格キー(スピードコントローラ)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                Case "3", "4"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-C-SC-" & strKSymbol
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-C-SC-" & strKSymbol
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
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-C-" & objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-C-" & objKtbnStrc.strcSelection.strOpSymbol(10).Trim
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
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-C-" & objKtbnStrc.strcSelection.strOpSymbol(11).Trim & CdCst.Sign.Hypen & strKSymbol
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "2"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NHS-C-" & objKtbnStrc.strcSelection.strOpSymbol(11).Trim & CdCst.Sign.Hypen & strKSymbol
                            decOpAmount(UBound(decOpAmount)) = 2
                    End Select
            End Select

            '(Z軸)オプション加算価格キー(アタッチメント)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "1"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NHS-C-ATATCHMENT-" & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "2"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NHS-C-ATATCHMENT-" & strXSymbol
                    decOpAmount(UBound(decOpAmount)) = 2
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
