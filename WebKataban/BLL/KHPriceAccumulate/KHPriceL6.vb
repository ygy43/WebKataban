'************************************************************************************
'*  ProgramID  ：KHPriceL6
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/12   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：省配線マニホールド　ベース配管タイプ　Ｍ(Ｄ)４ＳＢ１
'*             ：個別配線マニホールド　ベース配管タイプ　Ｍ(Ｄ)４ＳＢ１
'*             ：省配線マニホールド　タイレクト配管タイプ　Ｍ(Ｄ)３／４ＳＡ１
'*             ：個別配線マニホールド　タイレクト配管タイプ　Ｍ(Ｄ)３／４ＳＡ１
'*
'************************************************************************************
Module KHPriceL6

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim intLoopCnt As Integer
        Dim intStationQty As Integer = 0
        Dim intValveQtyD As Integer = 0
        Dim intValveQtyF As Integer = 0
        Dim intValveQtyS As Integer = 0

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            'バルブブロック連数
            intStationQty = CInt(objKtbnStrc.strcSelection.strOpSymbol(10).Trim)

            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                '仕様書形番が選択されていること、かつ、仕様書使用数が入っていること
                If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                   objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                        Case CdCst.Manifold.InspReportJp.Japanese, CdCst.Manifold.InspReportJp.English, _
                             CdCst.Manifold.InspReportEn.Japanese, CdCst.Manifold.InspReportEn.English
                            '加算なし
                        Case Else
                            Select Case intLoopCnt
                                Case 1 To 14
                                    '電磁弁
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-") - 1)
                                    decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.intQuantity(intLoopCnt))

                                    '電磁弁数(バルブ数)をカウントする
                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 1)
                                        Case "1"
                                            intValveQtyS = intValveQtyS + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Case Else
                                            intValveQtyD = intValveQtyD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                    End Select

                                    'A・Bポートフィルタ加算用に電磁弁をカウントする
                                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).IndexOf("F") >= 0 Or _
                                       objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).IndexOf("F") >= 0 Or _
                                       objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).IndexOf("F") >= 0 Then
                                        '電磁弁数(バルブ数)をカウントする
                                        intValveQtyF = intValveQtyF + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                Case 15 To 15
                                    'ベースマスキングプレート
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "4S1-MP"
                                    decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.intQuantity(intLoopCnt))

                                    'A・Bポートフィルタ加算用にベースマスキングプレートをカウントする
                                    'ベースマスキングプレートは継手CXのA・Bポートのみチェック）
                                    If objKtbnStrc.strcSelection.strCXAKataban(intLoopCnt).IndexOf("F") >= 0 Or _
                                       objKtbnStrc.strcSelection.strCXBKataban(intLoopCnt).IndexOf("F") >= 0 Then
                                        intValveQtyF = intValveQtyF + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                Case 16 To 19
                                    'ブランクプラグ,サイレンサ,ワンタッチ継手
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "4S1-" & objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.intQuantity(intLoopCnt))
                                Case 21 To 22
                                    'Case 22 To 23
                                    'ケーブル
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.intQuantity(intLoopCnt))
                            End Select
                    End Select
                End If
            Next

            'サブプレート加算価格キー
            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Case "SA1"
                    Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, 1)
                        Case "T"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "M*SA1-SPT-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "M*SA1-SP-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case "SB1"
                    Select Case Left(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, 1)
                        Case "T"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "M4SB1-SPT-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "M4SB1-SP-" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
            End Select

            '配線接続方式加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "" Then
                If Left(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, 1) = "T" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "4S1-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "4S1-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    decOpAmount(UBound(decOpAmount)) = CDec(intValveQtyS + intValveQtyD)
                End If
            End If

            '給排気ブロック加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "D" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "4S1-Q-D"
                If Left(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, 1) = "T" Then
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    decOpAmount(UBound(decOpAmount)) = 2
                End If
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "4S1-Q"
                If Left(objKtbnStrc.strcSelection.strOpSymbol(9).Trim, 1) = "T" Then
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    decOpAmount(UBound(decOpAmount)) = 2
                End If
            End If

            'DIN取付加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "D" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "4S1-DIN"
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'A・Bポートフィルタ加算価格キー
            If intValveQtyF <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "4S1-F"
                decOpAmount(UBound(decOpAmount)) = intValveQtyF
            End If

            '手動装置加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "4S1-M3"
                decOpAmount(UBound(decOpAmount)) = CDec(intValveQtyS + intValveQtyD)
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
