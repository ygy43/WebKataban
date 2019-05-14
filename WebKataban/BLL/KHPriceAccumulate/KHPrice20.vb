'************************************************************************************
'*  ProgramID  ：KHPrice20
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/05   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＦＲＬ　ＷＩＬＬ
'*
'************************************************************************************
Module KHPrice20

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt1 As Integer
        Dim intLoopCnt2 As Integer
        Dim intLoopCnt3 As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "3002E" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1)
                decOpAmount(UBound(decOpAmount)) = 1
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'オプション加算価格キー
            For intLoopCnt1 = 2 To objKtbnStrc.strcSelection.strOpSymbol.Length - 1
                If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt1).Trim = " " Then
                    Exit For
                End If

                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt1), CdCst.Sign.Delimiter.Comma)
                intLoopCnt3 = 0
                For intLoopCnt2 = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt2).Trim
                        Case ""
                        Case Else
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                Case "3000E", "B5102", "A7070", "A7019"
                                    Select Case strOpArray(intLoopCnt2).Trim
                                        Case "J", "EJ", "FJ", "F1J"
                                            intLoopCnt3 = 1
                                    End Select
                            End Select
                            Select Case strOpArray(intLoopCnt2).Trim
                                Case "Z", "M", "MG", "MG2"
                                    If intLoopCnt3 = 1 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt2).Trim & "J"
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt2).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                                Case Else
                                    If strOpArray(intLoopCnt2).Trim = "-G" Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                   strOpArray(intLoopCnt2).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   strOpArray(intLoopCnt2).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    End If
                            End Select
                    End Select
                Next
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
