'************************************************************************************
'*  ProgramID  ：KHPrice89
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/09   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：省配線マニホールド　ＬＭＦ０
'*
'************************************************************************************
Module KHPrice89

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
            Next

        Catch ex As Exception

            Throw ex

        End Try

        'Try

        '    '配列定義
        '    ReDim strOpRefKataban(0)
        '    ReDim decOpAmount(0)

        '    'ベース
        '    If Right(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 1) = "1" Then
        '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '        strOpRefKataban(UBound(strOpRefKataban)) = "LMF0-1-BASE-" & _
        '                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim
        '        decOpAmount(UBound(decOpAmount)) = 1
        '    Else
        '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '        strOpRefKataban(UBound(strOpRefKataban)) = "LMF0-2-BASE-" & _
        '                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim
        '        decOpAmount(UBound(decOpAmount)) = 1
        '    End If

        '    'A・Bポート接続口径
        '    If objKtbnStrc.strcSelection.intQuantity(9) > 0 Then
        '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '        strOpRefKataban(UBound(strOpRefKataban)) = "LMF0-PORT-C4"
        '        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(9)
        '    End If
        '    If objKtbnStrc.strcSelection.intQuantity(10) > 0 Then
        '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '        strOpRefKataban(UBound(strOpRefKataban)) = "LMF0-PORT-C6"
        '        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(10)
        '    End If
        '    If objKtbnStrc.strcSelection.intQuantity(11) > 0 Then
        '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '        strOpRefKataban(UBound(strOpRefKataban)) = "LMF0-PORT-01Z"
        '        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(11)
        '    End If

        '    'P・R1・R2ポート接続口径
        '    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
        '        Case "C8B", "C8D", "C8U"
        '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '            strOpRefKataban(UBound(strOpRefKataban)) = "LMF0-PORT-" & _
        '                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
        '            decOpAmount(UBound(decOpAmount)) = 1
        '    End Select

        '    '電気接続
        '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '    strOpRefKataban(UBound(strOpRefKataban)) = "LMF0-OP-" & _
        '                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
        '    decOpAmount(UBound(decOpAmount)) = 1

        'Catch ex As Exception

        '    Throw ex

        'End Try

    End Sub

End Module
