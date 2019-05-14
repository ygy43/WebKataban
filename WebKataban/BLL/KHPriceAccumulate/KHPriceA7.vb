'************************************************************************************
'*  ProgramID  ：KHPriceA7
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＮＷ３ＧＡ２
'*             ：ＮＷ４ＧＡ２
'*             ：ＮＷ４ＧＢ２
'*             ：ＮＷ４ＧＺ２
'*
'************************************************************************************
Module KHPriceA7

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
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "R1" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                'RM1805036_二次電池価格加算対応
                If objKtbnStrc.strcSelection.strKeyKataban = "P" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-P40"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

                '食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
                If objKtbnStrc.strcSelection.strKeyKataban = "F" Or objKtbnStrc.strcSelection.strKeyKataban = "H" Or objKtbnStrc.strcSelection.strKeyKataban = "O" Or objKtbnStrc.strcSelection.strKeyKataban = "L" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1

                    'バルブブロックキーの作成を追加  2017/03/07 追加
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-V-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1

                End If

            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                'RM1805036_二次電池価格加算対応
                If objKtbnStrc.strcSelection.strKeyKataban = "P" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-P40"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

                '食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
                If objKtbnStrc.strcSelection.strKeyKataban = "F" Or objKtbnStrc.strcSelection.strKeyKataban = "H" Or objKtbnStrc.strcSelection.strKeyKataban = "O" Or objKtbnStrc.strcSelection.strKeyKataban = "L" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1

                    'バルブブロックキーの作成を追加  2017/03/07 追加
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-V-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1

                End If
            End If

            'オプション価格加算キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "A", "F"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    Case "M7"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            Case "1", "11"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "W4G2" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & "S"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Case "2", "3", "4", "5"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "W4G2" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & "D"
                                decOpAmount(UBound(decOpAmount)) = 1

                        End Select

                        '食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
                    Case "M"
                        If objKtbnStrc.strcSelection.strKeyKataban = "F" Or objKtbnStrc.strcSelection.strKeyKataban = "H" Or objKtbnStrc.strcSelection.strKeyKataban = "O" Or objKtbnStrc.strcSelection.strKeyKataban = "L" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "W4G2" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                End Select
            Next

            '食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "4" Then
                If objKtbnStrc.strcSelection.strKeyKataban = "F" Or objKtbnStrc.strcSelection.strKeyKataban = "H" Or objKtbnStrc.strcSelection.strKeyKataban = "O" Or objKtbnStrc.strcSelection.strKeyKataban = "L" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-DC-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            '食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "1" Then
                If objKtbnStrc.strcSelection.strKeyKataban = "F" Or objKtbnStrc.strcSelection.strKeyKataban = "H" Or objKtbnStrc.strcSelection.strKeyKataban = "O" Or objKtbnStrc.strcSelection.strKeyKataban = "L" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-AC-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            'Hが含まれない場合は排気誤作動防止弁価格キー設定
            If objKtbnStrc.strcSelection.strOpSymbol(6).IndexOf("H") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & CdCst.Sign.Hypen & "H"
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            ''食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
            'If objKtbnStrc.strcSelection.strKeyKataban = "F" Or objKtbnStrc.strcSelection.strKeyKataban = "H" Or objKtbnStrc.strcSelection.strKeyKataban = "O" Then

            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '    strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-FP1"
            '    decOpAmount(UBound(decOpAmount)) = 1

            'End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
