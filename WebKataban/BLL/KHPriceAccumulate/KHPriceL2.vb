'************************************************************************************
'*  ProgramID  ：KHPriceL2
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/13   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：レギュレータブロックマニホールド　ＭＮＲＢ５００／ＭＮＲＪＢ５００
'*
'************************************************************************************
Module KHPriceL2

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim intLoopCnt As Integer

        Dim intStationQty As Integer

        'MNRB500用
        Dim intRB500_OptionUQty As Integer = 0
        Dim intRB500_OptionPQty As Integer = 0
        Dim intRB500_OptionLQty As Integer = 0
        Dim intRB500_OptionNQty As Integer = 0
        Dim intRB500_OptionTQty As Integer = 0
        Dim intRB500_OptionLTQty As Integer = 0
        Dim intRB500_OptionG39Qty As Integer = 0
        Dim intRB500_OptionMPQty As Integer = 0
        'MNRJB500用
        Dim intRJB500_OptionPQty As Integer = 0
        Dim intRJB500_OptionLQty As Integer = 0
        Dim intRJB500_OptionTQty As Integer = 0
        Dim intRJB500_OptionLTQty As Integer = 0
        Dim intRJB500_OptionMPQty As Integer = 0
        
        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'レギュレータブロック連数
            intStationQty = CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim)

            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                   objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                    Select Case intLoopCnt
                        Case 1
                            'エンドブロックL
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                        Case 2
                            '集中給気ブロック
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                        Case 3
                            'APS付集中給気ブロック
                            If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-3") = 0 And InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-5") = 0 Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 14)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6) & CdCst.Sign.Hypen & _
                                                                           Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 16, 1)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End If
                        Case 4 To 13
                            'レギュレータブロック
                            Select Case True
                                Case InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "4") <> 0
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "4"))
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "6") <> 0
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "6"))
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End Select

                            Select Case True
                                Case Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6) = "NRB500"
                                    'MNRB500用
                                    'オプション選択数カウント
                                    If InStr(12, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "U") <> 0 Then
                                        intRB500_OptionUQty = intRB500_OptionUQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    If InStr(12, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "P") <> 0 Then
                                        intRB500_OptionPQty = intRB500_OptionPQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    If InStr(12, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "L") <> 0 And _
                                        InStr(12, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "T") = 0 Then
                                        intRB500_OptionLQty = intRB500_OptionLQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    If InStr(12, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "N") <> 0 Then
                                        intRB500_OptionNQty = intRB500_OptionNQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    If InStr(12, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "T") <> 0 And _
                                        InStr(12, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "L") = 0 Then
                                        intRB500_OptionTQty = intRB500_OptionTQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    If InStr(12, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "L") <> 0 And _
                                        InStr(12, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "T") <> 0 Then
                                        intRB500_OptionLTQty = intRB500_OptionLTQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    If InStr(12, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "G39") <> 0 Then
                                        intRB500_OptionG39Qty = intRB500_OptionG39Qty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                Case Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 7) = "NRJB500"
                                    'MNRJB500用
                                    'オプション選択数カウント
                                    If InStr(13, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "P") <> 0 Then
                                        intRJB500_OptionPQty = intRJB500_OptionPQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    If InStr(13, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "L") <> 0 And _
                                        InStr(13, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "T") = 0 Then
                                        intRJB500_OptionLQty = intRJB500_OptionLQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    If InStr(13, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "T") <> 0 And _
                                        InStr(13, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "L") = 0 Then
                                        intRJB500_OptionTQty = intRJB500_OptionTQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    If InStr(13, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "L") <> 0 And _
                                        InStr(13, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "T") <> 0 Then
                                        intRJB500_OptionLTQty = intRJB500_OptionLTQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                            End Select
                        Case 14
                            'MP付サブベース
                            Select Case True
                                Case Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 7) = "NRB500A"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 14)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 7) = "NRB500B"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 15)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 8) = "NRJB500A"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 15)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 8) = "NRJB500B"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 16)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End Select

                            Select Case True
                                Case Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6) = "NRB500"
                                    'MNRB500用
                                    'オプション選択数カウント
                                    If InStr(12, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "U") <> 0 Then
                                        intRB500_OptionUQty = intRB500_OptionUQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    If InStr(12, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "MP") <> 0 Then
                                        intRB500_OptionMPQty = intRB500_OptionMPQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                Case Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 7) = "NRJB500"
                                    'MNRJB500用
                                    'オプション選択数カウント
                                    If InStr(13, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "MP") <> 0 Then
                                        intRJB500_OptionMPQty = intRJB500_OptionMPQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                            End Select
                        Case 15
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                        Case 16 To 18
                            'ブランクプラグ
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "NRB500-" & objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                    End Select
                End If
            Next

            '取付レール長さ
            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            strOpRefKataban(UBound(strOpRefKataban)) = "NRB500-BAA"
            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.decDinRailLength
            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail

            'MNRB500用
            'OUT側上配管(U)
            If intRB500_OptionUQty <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "RB500U"
                decOpAmount(UBound(decOpAmount)) = intRB500_OptionUQty
            End If

            'パネルマウント(P)
            If intRB500_OptionPQty <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "RB500P"
                decOpAmount(UBound(decOpAmount)) = intRB500_OptionPQty
            End If

            '低圧用(L)
            If intRB500_OptionLQty <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "RB500L"
                decOpAmount(UBound(decOpAmount)) = intRB500_OptionLQty
            End If

            'ノンリリーフ(N)
            If intRB500_OptionNQty <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "RB500N"
                decOpAmount(UBound(decOpAmount)) = intRB500_OptionNQty
            End If

            '圧力計なし(T)
            If intRB500_OptionTQty <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "RB500T"
                decOpAmount(UBound(decOpAmount)) = intRB500_OptionTQty
            End If

            '低圧用＋圧力計なし(LT)
            If intRB500_OptionLTQty <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "RB500LT"
                decOpAmount(UBound(decOpAmount)) = intRB500_OptionLTQty
            End If

            '圧力計(G39)
            If intRB500_OptionG39Qty <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "RB500G39"
                decOpAmount(UBound(decOpAmount)) = intRB500_OptionG39Qty
            End If

            'マスキングプレート付(MP)
            If intRB500_OptionMPQty <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "RB500MP"
                decOpAmount(UBound(decOpAmount)) = intRB500_OptionMPQty
            End If

            'MNRJB500用
            'パネルマウント(P)
            If intRJB500_OptionPQty <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "RJB500P"
                decOpAmount(UBound(decOpAmount)) = intRJB500_OptionPQty
            End If

            '低圧用(L)
            If intRJB500_OptionLQty <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "RJB500L"
                decOpAmount(UBound(decOpAmount)) = intRJB500_OptionLQty
            End If

            '圧力計なし(T)
            If intRJB500_OptionTQty <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "RJB500T"
                decOpAmount(UBound(decOpAmount)) = intRJB500_OptionTQty
            End If

            '低圧用＋圧力計なし(LT)
            If intRJB500_OptionLTQty <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "RJB500LT"
                decOpAmount(UBound(decOpAmount)) = intRJB500_OptionLTQty
            End If

            'マスキングプレート付(MP)
            If intRJB500_OptionMPQty <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "RJB500MP"
                decOpAmount(UBound(decOpAmount)) = intRJB500_OptionMPQty
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
