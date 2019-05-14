'************************************************************************************
'*  ProgramID  ：KHPriceP0
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/05/22   作成者：NII A.Takahashi
'*
'*  概要       ：ラピフロー   ＦＳＭ２・ＦＳＭ３シリーズ
'*
'************************************************************************************
Module KHPriceP0

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'RM1802023_FSM3シリーズ追加
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim

                Case "FSM3"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim

                    decOpAmount(UBound(decOpAmount)) = 1

                    'バルブオプション加算価格キー
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4", "5", "6"  'ステンレスボディ
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim & "-SUS"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else           '樹脂
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'ケーブル加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(10).Trim.Length <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '取付アタッチメント加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(11).Trim.Length <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '添付書類加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(12).Trim.Length <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'クリーン仕様加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(13).Trim.Length <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "FSM2"

                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim.Length <> 0 Then
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "*" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                            Else
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "*" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            End If
                            decOpAmount(UBound(decOpAmount)) = 1

                            'ケーブル加算価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(1).Trim)
                                    Case "N", "P"
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                  "N/P" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                    Case "A"
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   "A" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                End Select
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            'ブラケット加算価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            'トレーサビリティ加算価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            'ニードル弁付き加算価格キー
                            '例)"FSM2-N-U2L-H","FSM2-N-O5L-H","FSM2-N-S"
                            If objKtbnStrc.strcSelection.strOpSymbol(10).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                '接続口径(ボディ材質)判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    'ボディが樹脂材質の場合
                                    Case "H04", "H06", "H08", "H10", "H08"
                                        '流量レンジ判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            '２リットル／秒以下の場合
                                            Case "005", "010", "020"
                                                strOpRefKataban(UBound(strOpRefKataban)) = "FSM2" & "-N-U2L-H"
                                                '５リットル／秒以上の場合
                                            Case Else
                                                strOpRefKataban(UBound(strOpRefKataban)) = "FSM2" & "-N-O5L-H"
                                        End Select
                                        'ボディがステンレス材質の場合
                                    Case "S06", "S08"
                                        strOpRefKataban(UBound(strOpRefKataban)) = "FSM2" & "-N-S"
                                    Case Else
                                End Select

                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            'クリーン仕様加算価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(11).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "FSM2" & "-" & objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case "D"
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)

                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'ケーブル加算価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(2).Trim)
                                    Case "N", "P"
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                  "N/P" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "A"
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                   "A" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                End Select
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            'ブラケット加算価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            'クリーン仕様加算価格キー
                            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim.Length <> 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "FSM2" & "-D-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
