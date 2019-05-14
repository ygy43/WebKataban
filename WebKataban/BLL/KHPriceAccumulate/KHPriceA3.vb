'************************************************************************************
'*  ProgramID  ：KHPriceA3
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/05   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ギャップスイッチユニット　ＵＧＰＳ２
'*
'************************************************************************************
Module KHPriceA3

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                Case "C0", "C1", "C3", "C5"
                    'コネクタ形ユニット
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "C" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(10).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "F"
                    'DIN端子ユニット
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "F" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(10).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "CTL", "CTR"
                    'コネクタ形集中端子ユニット
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "CT" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(10).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "TL", "TR", "T1", "T2", "T3", "T4"
                    'リード線形集中端子ユニット
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "T" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(10).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '配線オプション加算価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(6).Trim
            decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(3).Trim)

            'アタッチメント加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "L"
                                decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(3).Trim)
                            Case Else
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select
            Next

            '圧力計加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(3).Trim)
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
