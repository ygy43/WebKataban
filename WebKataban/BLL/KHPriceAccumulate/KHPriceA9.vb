'************************************************************************************
'*  ProgramID  ：KHPriceA9
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/05   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ベース搭載用　電磁弁単品　Ｗ３ＧＡ２／Ｗ４ＧＡ２／Ｗ４ＧＢ２
'*
'************************************************************************************
Module KHPriceA9

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim bolOptionH As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
            decOpAmount(UBound(decOpAmount)) = 1

            'RM1805036_二次電池価格加算対応
            If objKtbnStrc.strcSelection.strKeyKataban = "G" Or objKtbnStrc.strcSelection.strKeyKataban = "U" Then
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-P40"
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
            If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "-FP1"
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'オプション検索
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "H"
                        bolOptionH = True
                End Select
            Next

            '排気誤動作防止弁付の減算
            If bolOptionH = False Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "H"
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "A", "F"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
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
                        If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "W4G2" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                End Select
            Next

            '食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "4" Then
                If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-DC-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

            '食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "1" Then
                If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-AC-FP1"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
