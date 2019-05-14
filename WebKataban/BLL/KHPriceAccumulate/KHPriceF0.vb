'************************************************************************************
'*  ProgramID  ：KHPriceF0
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/26   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ニューハンドリングシステム　ＮＳＲ
'*
'************************************************************************************
Module KHPriceF0

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim intStroke As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'ストローク取得
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "1"
                    intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                          CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim), _
                                                          CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim))
                Case "2"
                    intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                          CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim), _
                                                          CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim))
            End Select

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                       intStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1

            'オプション加算価格キー(2ヘッド)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "2"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-2HEAD-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション加算価格キー(ストローク調整ブロック)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                Case "L", "R", "D"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-STAB-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション加算価格キー(スピードコントローラ・取付ブラケット)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                Case "3", "4", "5", "6", "7", "8"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-SCLB-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション加算価格キー(ケーブルベア)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                Case "B", "F", "W", "X"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-CABLE-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション加算価格キー(スイッチ)
            Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                Case "A"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
