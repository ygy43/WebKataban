'************************************************************************************
'*  ProgramID  ：KHPriceB4
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/12   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：省配線ブロックマニホールド　ＭＮ４ＴＢ１／ＭＮ４ＴＢ２／ＭＮ４ＴＢＸ１２
'*
'************************************************************************************
Module KHPriceB4

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim bolOptionK As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "R1" Then
                '省配線
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-MBASE"
                decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
            Else
                '個別配線
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-R1-MBASE"
                decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
            End If

            ''食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
            'If objKtbnStrc.strcSelection.strKeyKataban = "F" Then

            '    If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "FP1" Then
            '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '        strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-FP1"
            '        decOpAmount(UBound(decOpAmount)) = 1
            '    Else
            '        '合致しない場合は価格キーを作成しない
            '    End If

            'End If

            '電装ブロック加算価格キー(省配線のときのみ)
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "R1" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2" & "-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                ''集中端子台の場合のみ以下の処理を実施する  RM1702019  2017/03/01 追加
                If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2" & "-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim & "-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            '給排気ブロック加算価格キー
            '外部パイロット選択有無の判定
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "K"
                        bolOptionK = True
                End Select
            Next

            Select Case bolOptionK
                Case True
                    '排気方法判定
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2" & "-QK-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        '食品製造工程向けオプション追加に伴う処理の追加  2017/03/01 追加
                        If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2" & "-QK-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-FP1"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2" & "-QK-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        '食品製造工程向けオプション追加に伴う処理の追加  2017/03/01 追加
                        If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2" & "-QK-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-FP1"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                Case False
                    '排気方法判定
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2" & "-Q-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        '食品製造工程向けオプション追加に伴う処理の追加  2017/03/01 追加
                        If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2" & "-Q-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-FP1"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2" & "-Q-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        '食品製造工程向けオプション追加に伴う処理の追加  2017/03/01 追加
                        If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2" & "-Q-" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-FP1"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
            End Select

            'エンドブロック加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "R1" Then
                '省配線(排気方法判定)
                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-ER"
                    decOpAmount(UBound(decOpAmount)) = 1

                    '食品製造工程向けオプション追加に伴う処理の追加  2017/03/01 追加
                    If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-ER-FP1"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-EXR"
                    decOpAmount(UBound(decOpAmount)) = 1

                    '食品製造工程向けオプション追加に伴う処理の追加  2017/03/01 追加
                    If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-EXR-FP1"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                End If
            Else
                '個別配線
                If InStr(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, "X") = 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-ER"
                    decOpAmount(UBound(decOpAmount)) = 1

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-EL"
                    decOpAmount(UBound(decOpAmount)) = 1

                    '食品製造工程向けオプション追加に伴う処理の追加  2017/03/01 追加
                    If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-ER-FP1"
                        decOpAmount(UBound(decOpAmount)) = 1

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-EL-FP1"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Else
                    If InStr(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, "D") <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-ER"
                        decOpAmount(UBound(decOpAmount)) = 1

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-EXL"
                        decOpAmount(UBound(decOpAmount)) = 1

                        '食品製造工程向けオプション追加に伴う処理の追加  2017/03/01 追加
                        If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-ER-FP1"
                            decOpAmount(UBound(decOpAmount)) = 1

                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-EXL-FP1"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-EXR"
                        decOpAmount(UBound(decOpAmount)) = 1

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-EL"
                        decOpAmount(UBound(decOpAmount)) = 1

                        '食品製造工程向けオプション追加に伴う処理の追加  2017/03/01 追加
                        If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-EXR-FP1"
                            decOpAmount(UBound(decOpAmount)) = 1

                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-EL-FP1"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                End If
            End If

            ''オプション加算価格キー
            'strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
            'For intLoopCnt = 0 To strOpArray.Length - 1
            '    Select Case strOpArray(intLoopCnt).Trim
            '        Case "F"
            '            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban, 2, 5) & CdCst.Sign.Hypen & _
            '                                                       strOpArray(intLoopCnt).Trim
            '            decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
            '    End Select
            'Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
