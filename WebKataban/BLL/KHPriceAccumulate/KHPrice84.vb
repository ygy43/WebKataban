'************************************************************************************
'*  ProgramID  ：KHPrice84
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/06   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：小形ピストンタイプレギュレータ　ＲＡ８００
'*             ：小形フィルタレギュレータ　ＷＢ５００
'*
'************************************************************************************
Module KHPrice84

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt1 As Integer
        Dim intLoopCnt2 As Integer

        Dim bolOptionL As Boolean = False
        Dim bolOptionT As Boolean = False
        Dim bolOptionP As Boolean = False
        Dim bolSkip As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '選択されたオプション情報を取得する
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt1 = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt1).Trim
                    Case "L"
                        bolOptionL = True
                    Case "T"
                        bolOptionT = True
                    Case "P"
                        bolOptionP = True
                End Select
            Next

            '基本価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "WB500"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "RA800"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション加算価格キー
            For intLoopCnt1 = 4 To 5
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt1), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt2 = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt2).Trim
                        Case ""
                        Case "L", "T"
                            If bolOptionL = True And bolOptionT = True Then
                                If bolSkip = True Then
                                Else
                                    bolSkip = True

                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "LT"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           strOpArray(intLoopCnt2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "P"
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                Case "RA800"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                               strOpArray(intLoopCnt2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "WB500"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                               strOpArray(intLoopCnt2).Trim
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       strOpArray(intLoopCnt2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Next
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
