'************************************************************************************
'*  ProgramID  ：KHPriceJ0
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/07   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ハイクール
'*             ：ＨＹＷ２０００／２０００Ｃ／３０００／３０００Ｗ／５０００／６０００
'*
'************************************************************************************
Module KHPriceJ0

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim bolOptionX As Boolean = False
        Dim bolOptionP7 As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                '基本価格キー
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                'オプション加算価格キー
                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    Case "HYW10", "HYW20", "HYW30", "HYW50"
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(2), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case ""
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "00" & CdCst.Sign.Hypen & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Next
                End Select
            Else
                '基本価格キー
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                'オプション加算価格キー
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "00" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Next
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
