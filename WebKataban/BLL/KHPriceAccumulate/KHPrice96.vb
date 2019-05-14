'************************************************************************************
'*  ProgramID  ：KHPrice96
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/13   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＰＶ５形マニホールド　ＣＭＦ
'*
'************************************************************************************
Module KHPrice96

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

        '    '基本価格キー
        '    If Left(objKtbnStrc.strcSelection.strFullKataban.Trim, 4) = "CMFZ" Then
        '        'マニホールドブロック
        '        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
        '            Case "HX3"
        '                'CMF1
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE1-BLOCK-02"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Mid(objKtbnStrc.strcSelection.strFullKataban.Trim, Len(objKtbnStrc.strcSelection.strFullKataban.Trim) - 1, 1))

        '                'CMF2
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE2-BLOCK-03"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Right(objKtbnStrc.strcSelection.strFullKataban.Trim, 1))
        '            Case "HX4"
        '                'CMF1
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE1-BLOCK-02"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Mid(objKtbnStrc.strcSelection.strFullKataban.Trim, Len(objKtbnStrc.strcSelection.strFullKataban.Trim) - 1, 1))

        '                'CMF2
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE2-BLOCK-04"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Right(objKtbnStrc.strcSelection.strFullKataban.Trim, 1))
        '            Case "HX5"
        '                'CMF1
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE1-BLOCK-03"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Mid(objKtbnStrc.strcSelection.strFullKataban.Trim, Len(objKtbnStrc.strcSelection.strFullKataban.Trim) - 1, 1))

        '                'CMF2
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE2-BLOCK-03"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Right(objKtbnStrc.strcSelection.strFullKataban.Trim, 1))
        '            Case "HX6"
        '                'CMF1
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE1-BLOCK-03"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Mid(objKtbnStrc.strcSelection.strFullKataban.Trim, Len(objKtbnStrc.strcSelection.strFullKataban.Trim) - 1, 1))

        '                'CMF2
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE2-BLOCK-04"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Right(objKtbnStrc.strcSelection.strFullKataban.Trim, 1))
        '        End Select
        '        'フート
        '        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
        '            Case "HY3"
        '                'CMF1
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE1-FOOT-03"
        '                decOpAmount(UBound(decOpAmount)) = 1

        '                'CMF2
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE2-FOOT-04"
        '                decOpAmount(UBound(decOpAmount)) = 1
        '            Case "HY4"
        '                'CMF1
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE1-FOOT-03"
        '                decOpAmount(UBound(decOpAmount)) = 1

        '                'CMF2
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE2-FOOT-06"
        '                decOpAmount(UBound(decOpAmount)) = 1
        '            Case "HY5"
        '                'CMF1
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE1-FOOT-04"
        '                decOpAmount(UBound(decOpAmount)) = 1

        '                'CMF2
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE2-FOOT-04"
        '                decOpAmount(UBound(decOpAmount)) = 1
        '            Case "HY6"
        '                'CMF1
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE1-FOOT-04"
        '                decOpAmount(UBound(decOpAmount)) = 1

        '                'CMF2
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMFZ-BASE2-FOOT-06"
        '                decOpAmount(UBound(decOpAmount)) = 1
        '        End Select

        '        'ミックスブロック
        '        For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
        '            Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
        '                Case "CMFBZ-00L", "CMFBZ-00R"
        '                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
        '                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
        '            End Select
        '        Next
        '    Else
        '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strFullKataban.Trim, 1, InStr(1, objKtbnStrc.strcSelection.strFullKataban.Trim, "-")) & "MANIHOLD-BASE"
        '        decOpAmount(UBound(decOpAmount)) = 1
        '    End If

        '    'A・Bポート口径価格キー
        '    '裏配管の場合
        '    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "Z" Then
        '        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
        '            Case "HX1"
        '                'CMF1
        '                '02加算
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strFullKataban.Trim, 4) & "-BASE-PORT-02Z"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Mid(objKtbnStrc.strcSelection.strFullKataban.Trim, Len(objKtbnStrc.strcSelection.strFullKataban.Trim) - 1, 1))

        '                '03加算
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strFullKataban.Trim, 4) & "-BASE-PORT-03Z"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Right(objKtbnStrc.strcSelection.strFullKataban.Trim, 1))
        '            Case "HX2"
        '                'CMF2
        '                '03加算
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strFullKataban.Trim, 4) & "-BASE-PORT-03Z"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Mid(objKtbnStrc.strcSelection.strFullKataban.Trim, Len(objKtbnStrc.strcSelection.strFullKataban.Trim) - 1, 1))

        '                '04加算
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strFullKataban.Trim, 4) & "-BASE-PORT-04Z"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Right(objKtbnStrc.strcSelection.strFullKataban.Trim, 1))
        '            Case "HX3"
        '                'CMFZ
        '                '02加算
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMF1-BASE-PORT-02Z"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Mid(objKtbnStrc.strcSelection.strFullKataban.Trim, Len(objKtbnStrc.strcSelection.strFullKataban.Trim) - 1, 1))

        '                '03加算
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMF2-BASE-PORT-03Z"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Right(objKtbnStrc.strcSelection.strFullKataban.Trim, 1))
        '            Case "HX4"
        '                'CMFZ
        '                '02加算
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMF1-BASE-PORT-02Z"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Mid(objKtbnStrc.strcSelection.strFullKataban.Trim, Len(objKtbnStrc.strcSelection.strFullKataban.Trim) - 1, 1))

        '                '04加算
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMF2-BASE-PORT-04Z"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Right(objKtbnStrc.strcSelection.strFullKataban.Trim, 1))
        '            Case "HX5"
        '                'CMFZ
        '                '03加算
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMF1-BASE-PORT-03Z"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Mid(objKtbnStrc.strcSelection.strFullKataban.Trim, Len(objKtbnStrc.strcSelection.strFullKataban.Trim) - 1, 1))

        '                '03加算
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMF2-BASE-PORT-03Z"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Right(objKtbnStrc.strcSelection.strFullKataban.Trim, 1))
        '            Case "HX6"
        '                'CMFZ
        '                '03加算
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMF1-BASE-PORT-03Z"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Mid(objKtbnStrc.strcSelection.strFullKataban.Trim, Len(objKtbnStrc.strcSelection.strFullKataban.Trim) - 1, 1))

        '                '04加算
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMF2-BASE-PORT-04Z"
        '                decOpAmount(UBound(decOpAmount)) = CDec(Right(objKtbnStrc.strcSelection.strFullKataban.Trim, 1))
        '            Case Else
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strFullKataban.Trim, 4) & "-BASE-PORT-" & _
        '                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
        '                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
        '                decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(2).Trim)
        '        End Select
        '    End If

        '    'P・Rポート口径価格キー
        '    'P・Rポートの加算は"06"または"HY2"の場合のみで、"06"・"HY2"はCMF2の場合にしかありえない
        '    If Left(objKtbnStrc.strcSelection.strFullKataban.Trim, 4) = "CMF2" Then
        '        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
        '            Case "04"
        '            Case Else
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strFullKataban.Trim, 4) & "-BASE-PORT-" & _
        '                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
        '                decOpAmount(UBound(decOpAmount)) = 1
        '        End Select
        '    End If

        '    '制御ユニット価格キー
        '    '制御ユニット付のものは加算
        '    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
        '        Case "4", "5", "9"
        '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '            strOpRefKataban(UBound(strOpRefKataban)) = "CMF1-UNIT-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
        '            decOpAmount(UBound(decOpAmount)) = 1
        '        Case "6", "7"
        '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '            strOpRefKataban(UBound(strOpRefKataban)) = "CMF1-BASE-UNIT-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
        '            decOpAmount(UBound(decOpAmount)) = 1
        '    End Select

        '    'サイレンサボックス価格キー
        '    '制御ユニット付でないものは、サイレンサボックスが選択可能
        '    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
        '        Case "1", "2", "3", "8"
        '            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
        '                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
        '                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
        '                strOpRefKataban(UBound(strOpRefKataban)) = "CMF-BASE-" & objKtbnStrc.strcSelection.strOpSymbol(8).Trim
        '                decOpAmount(UBound(decOpAmount)) = 1
        '            End If
        '    End Select

        'Catch ex As Exception

        '    Throw ex

        'End Try

    End Sub

End Module
