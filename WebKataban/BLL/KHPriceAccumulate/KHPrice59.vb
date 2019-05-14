'************************************************************************************
'*  ProgramID  ：KHPrice59
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/12   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：省配線バルブ　ＭＮ３Ｓ０／ＭＮ４Ｓ０／ＭＴ３Ｓ０／ＭＴ４Ｓ０
'*
'************************************************************************************
Module KHPrice59

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim intLoopCnt As Integer
        Dim intStationQty As Integer = 0
        Dim intQuantity As Integer = 0
        Dim IntValveQty3P As Integer = 0
        Dim IntValveQty4P As Integer = 0

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'バルブブロック連数
            intStationQty = CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)

            '予備ケーブル専用(MN3S0,MN4S0専用)
            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "MN3S0" Or _
               objKtbnStrc.strcSelection.strSeriesKataban.Trim = "MN4S0" Then
                If objKtbnStrc.strcSelection.strOptionKataban(22).Trim <> "" Then
                    objKtbnStrc.strcSelection.intQuantity(22) = 1
                Else
                    objKtbnStrc.strcSelection.intQuantity(22) = 0
                End If
            End If

            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                   objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                        Case CdCst.Manifold.InspReportJp.Japanese, CdCst.Manifold.InspReportJp.English, _
                             CdCst.Manifold.InspReportEn.Japanese, CdCst.Manifold.InspReportEn.English
                            '加算なし
                        Case Else
                            Select Case intLoopCnt
                                Case 2 To 8
                                    'バルブブロック
                                    '手動装置"-M1"除去
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & " ", "-M1 ") <> 0 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & " ", InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & " ", "-M1 ") - 1)
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        '手動装置"-M2"除去
                                        If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & " ", "-M2 ") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & " ", InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & " ", "-M2 ") - 1)
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            '手動装置"-M3"除去
                                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & " ", "-M3 ") <> 0 Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & " ", InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & " ", "-M3 ") - 1)
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Else
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End If
                                        End If
                                    End If
                                    'A・Bポートフィルター"F"除去
                                    If InStr(1, strOpRefKataban(UBound(strOpRefKataban)) & " ", "F ") <> 0 Then
                                        strOpRefKataban(UBound(strOpRefKataban)) = Left(strOpRefKataban(UBound(strOpRefKataban)) & " ", InStr(1, strOpRefKataban(UBound(strOpRefKataban)) & " ", "F ") - 1)
                                    End If

                                    '電磁弁数(バルブ数)をカウントする
                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 1)
                                        Case "3"
                                            IntValveQty3P = IntValveQty3P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Case "4"
                                            IntValveQty4P = IntValveQty4P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End Select

                                Case 1, 9 To 14
                                    'Case 1, 9 To 14, 20
                                    '配線ブロック,給排気ブロック,仕切りブロック,エンドブロック,取付レール
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case 15 To 19
                                    'サイレンサ,ブランクプラグ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "N4S0" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case 21
                                    'ケーブル
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case 22
                                    'ケーブル
                                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                        Case "MN3S0", "MN4S0"
                                            If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.StartsWith("1") Then
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = "N4T-SUBCABLE-(1-2)"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Else
                                                If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).StartsWith("3") Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = "N4T-SUBCABLE-(3-6)"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                Else
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End If
                                            End If
                                        Case Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End Select
                            End Select

                            Select Case Left(strOpRefKataban(UBound(strOpRefKataban)), 8)
                                Case "N4S0-GZP"
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Joint
                            End Select

                            'バルブブロックの時
                            If (Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) = "N4S0" Or _
                                Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 4) = "N3S0") And _
                                Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 1) >= "0" And _
                                Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 1) <= "9" Then
                                If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 1) = 1 Then
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                End If
                            End If
                    End Select
                End If
            Next

            'レール長さ加算価格キー
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = "N4S0-BAA"
            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.decDinRailLength
            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail

            'A・Bポートフィルタ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                If IntValveQty3P <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "N3S0" & "-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = IntValveQty3P
                End If
                If IntValveQty4P <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "N4S0" & "-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = IntValveQty4P
                End If
            End If

            '手動装置加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "N4S0" & "-" & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                decOpAmount(UBound(decOpAmount)) = intQuantity
            End If

            '配線方式(個別コネクタ)加算価格キー
            If Mid(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 1, 1) = "C" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "N4S0" & "-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                decOpAmount(UBound(decOpAmount)) = intStationQty
            End If

            'ダイレクトマウント加算価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "MT3S0", "MT4S0"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "MT4S0-DM"
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
