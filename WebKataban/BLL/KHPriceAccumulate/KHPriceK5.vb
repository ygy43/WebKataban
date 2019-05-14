'************************************************************************************
'*  ProgramID  ：KHPriceK5
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/27   作成者：NII K.Sudoh
'*
'*  概要       ：フリーポジション落下防止付セレックスシリンダ　ＵＳＣ
'*
'*  更新履歴   ：                       更新日：2007/05/16   更新者：NII A.Takahashi
'*               ・T2W/T3Wスイッチ追加に伴い、リード線加算ロジック部を修正
'************************************************************************************
Module KHPriceK5

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim intStroke As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'ストローク取得
            intStroke = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim), _
                                                  CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim))

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                       intStroke.ToString
            decOpAmount(UBound(decOpAmount)) = 1

            '支持形式加算価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            'スイッチ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                    Case "A", "B"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)

                        'リード線長さ加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                            If objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("L2") >= 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim & "-L2-" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                            Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "T2YD"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(2)-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                                    Case "T2YDT"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(3)-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                                    Case "T1H", "T1V", "T2H", "T2V", "T2YH", "T2YV", "T3H", "T3V", "T3YH", "T3YV", _
                                         "T0H", "T0V", "T5H", "T5V", "T8H", "T8V", "T2JH", "T2JV", "T2WH", "T2WV", "T3WH", "T3WV"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(4)-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                                    Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(5)-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(1)-" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                        decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                                End Select

                            End If
                        End If
                End Select
            End If

            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(9), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "N", "S", "T", "P6"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                   intStroke.ToString
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '付属品加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            'L2(耐強磁界スイッチ)加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("L2") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-L2-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'G1(コイルスクレーパ)加算価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.IndexOf("G1") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-G1-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
