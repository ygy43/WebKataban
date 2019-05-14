'************************************************************************************
'*  ProgramID  ：KHPrice12
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/13   作成者：NII K.Sudoh
'*
'*  概要       ：マイクロシリンダ　ＣＭＡ２シリーズ
'*
'*  更新履歴   ：                       更新日：2007/07/06   更新者：NII A.Takahashi
'*               ・T形スイッチ追加に伴い、リード線加算ロジック部を修正
'************************************************************************************
Module KHPrice12

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer

        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'ストローク取得
            intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim), _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim))

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                       intStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1

            '支持形式加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "CA", "CB", "TA", "TB"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'スイッチ加算価格キー
            If Right(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) <> "T" Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    Case "", "A", "B"
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "T0H", "T0V", "T2H", "T2V", "T3H", "T3V", "T5H", "T5V", "T1H", "T1V", "T8H", "T8V", _
                                 "T2WH", "T2WV", "T3WH", "T3WV", "T2YH", "T2YV", "T3YH", "T3YV", "T3PH", "T3PV"
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                           Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 1) & "(2)"
                            Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV"
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                           Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 1) & "(3)"
                            Case "T2JH", "T2JV"
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                           Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 1) & "(4)"
                            Case Else
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                           Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 1)
                        End Select
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                End Select

                'リード線長さ加算価格キー
                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    Case "A", "B"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim & _
                                                                   Left(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, 1)
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                End Select
            End If

            'オプション加算価格キー
            If Right(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "T" Then
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
            Else
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
            End If
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "M0", "M1"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                           strOpArray(intLoopCnt).Trim
                            Case "M"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                           intStroke.ToString
                            Case "N", "P"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           intStroke.ToString
                        End Select

                        If Right(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "D" Then
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "J", "K", "L", "M", "N", "F", "FE"
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Case Else
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Else
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                End Select
            Next

            '付属品加算価格キー
            If Right(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "T" Then
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(5), CdCst.Sign.Delimiter.Comma)
            Else
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
            End If
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "I"
                        bolOptionI = True
                    Case "Y"
                        bolOptionY = True
                End Select
            Next

            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        If Right(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "D" Then
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "I"
                                    If bolOptionY = True Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    End If
                                Case "Y"
                                    If bolOptionI = True Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    End If
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                End Select
            Next

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
