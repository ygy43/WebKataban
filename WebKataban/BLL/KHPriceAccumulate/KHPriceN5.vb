'************************************************************************************
'*  ProgramID  ：KHPriceN5
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/12   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：省配線ブロックマニホールド(Ｔ１／Ｔ６シリーズ)　ＭＷ４ＧＢ４
'*             ：個別配線ブロックマニホールド(Ｉ／Ｏコネクタタイプ)　ＭＷ４ＧＢ４
'*             ：省配線ブロックマニホールド(Ｔ１／Ｔ６シリーズ)　ＭＷ４ＧＺ４
'*             ：個別配線ブロックマニホールド(Ｉ／Ｏコネクタタイプ)　ＭＷ４ＧＺ４
'*
'*  更新履歴
'*                                      更新日：2008/04/09   更新者：T.Sato
'*　・受付No：RM0803048  オプションに『無記号』を追加したので価格キー作成ロジックを追加
'*                                      更新日：2008/07/20   更新者：T.Sato
'*  ・受付No：RM0805028　仕切りブロック追加に伴う変更
'*
'************************************************************************************
Module KHPriceN5

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim intQuantity As Integer = 0
        Dim intStationQty As Integer = 0
        Dim intValveQty1SWD As Integer = 0
        Dim intValveQty5SWD As Integer = 0
        Dim intScrewQty As Integer = 0
        Dim bolScrewFlg As Boolean = False
        Dim bolOptionK As Boolean = False
        Dim strRensu As String = String.Empty
        Dim strOption As String = String.Empty
        Dim strVoltage As String = String.Empty

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "S", "Y"
                    strRensu = objKtbnStrc.strcSelection.strOpSymbol(9)
                    strOption = objKtbnStrc.strcSelection.strOpSymbol(8)
                    strVoltage = objKtbnStrc.strcSelection.strOpSymbol(10)
                Case Else
                    strRensu = objKtbnStrc.strcSelection.strOpSymbol(8)
                    strOption = objKtbnStrc.strcSelection.strOpSymbol(7)
                    strVoltage = objKtbnStrc.strcSelection.strOpSymbol(9)
            End Select
           
            'バルブブロック連数
            intStationQty = CDec(strRensu.Trim)

            'ねじ
            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Case "08G", "08N", "10G", "10N"
                    bolScrewFlg = True
                Case Else
                    bolScrewFlg = False
            End Select

            '外部パイロット選択チェック
            strOpArray = Split(strOption, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "K"
                        bolOptionK = True
                End Select
            Next

            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                   objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                        Case CdCst.Manifold.InspReportJp.Japanese, CdCst.Manifold.InspReportJp.English, _
                             CdCst.Manifold.InspReportEn.Japanese, CdCst.Manifold.InspReportEn.English
                            '加算なし
                        Case Else
                            Select Case intLoopCnt
                                Case 1, 2
                                    '入出力ブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 8, Len(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim))
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case 3 To 4
                                    'エンドブロック
                                    If bolScrewFlg Then
                                        If bolOptionK Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Length - 2)
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Length - 1)
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If

                                        'ねじをカウントする
                                        intScrewQty = intScrewQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                Case 5
                                    '配線ブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case 6 To 10
                                    '電磁弁付バルブブロック
                                    If bolScrewFlg Then
                                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "R1" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Length - 1)
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Length - 1) & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If

                                        'ねじをカウントする
                                        intScrewQty = intScrewQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "R1" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If

                                        'RM1805036_二次電池加算価格対応
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            Case "C8", "C10", "C12", "CX"
                                                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                    Case "P", "Y"
                                                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "R1" Then
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-P40"
                                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        Else
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & "-P40"
                                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        End If
                                                End Select
                                        End Select
                                    End If

                                    '電磁弁数(バルブ数)をカウントする③
                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 7, 1)
                                        Case "1"
                                            intValveQty1SWD = intValveQty1SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Case "2", "3", "4", "5"
                                            intValveQty5SWD = intValveQty5SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End Select
                                Case 11 To 12
                                    'MPV付バルブブロック
                                    If bolScrewFlg Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Length - 1)
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                        'ねじをカウントする
                                        intScrewQty = intScrewQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                        'RM1805036_二次電池加算価格対応
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            Case "C8", "C10", "C12", "CX"
                                                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                    Case "P", "Y"
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-P40"
                                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                End Select
                                        End Select
                                    End If
                                Case 13
                                    '仕切りブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case 14 To 18
                                    'RM1210067 2013/02/12 Y.Tachi
                                    'Sレギュレータ(P,A,B),単独給・排気スペーサ
                                    'Select Case Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1)
                                    '    Case "G", "N"
                                    '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    '        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    '        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Length - 1)
                                    '        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    '        'ねじをカウントする
                                    '        intScrewQty = intScrewQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    '    Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    'End Select
                                Case 19
                                    '仕切プラグ(P)
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-P"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case 20
                                    '仕切プラグ(R)
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-R"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                    'decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case 21 To 23
                                    'ブランクプラグ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "W4G4-" & objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Case 24
                                    'ケーブルクランプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End Select

                            Select Case intLoopCnt
                                Case 4 To 8
                                    'AC110Vの時、電圧加算
                                    If strVoltage.Trim = "5" Then
                                        If InStr(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "410") <> 0 Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "W4G4-AC"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "W4G4-AC(2)"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End If
                                    End If
                            End Select

                            '電磁弁バルブブロックの時
                            If (Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) = "M4TB3" Or _
                                Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5) = "M4TB4") And _
                                Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) >= "0" And _
                                Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) <= "9" Then
                                If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 6, 1) = 1 Then
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                End If
                            End If
                    End Select
                End If
            Next

            'オプション加算価格キー(M/M7)
            Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                Case ""
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "W4G4-" & "BLANK" & "-S"
                    decOpAmount(UBound(decOpAmount)) = intValveQty1SWD

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "W4G4-" & "BLANK" & "-D"
                    decOpAmount(UBound(decOpAmount)) = intValveQty5SWD

                Case "M7"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "W4G4-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim & "-S"
                    decOpAmount(UBound(decOpAmount)) = intValveQty1SWD

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "W4G4-" & objKtbnStrc.strcSelection.strOpSymbol(6).Trim & "-D"
                    decOpAmount(UBound(decOpAmount)) = intValveQty5SWD

            End Select

            'オプション加算価格キー(A/K)
            strOpArray = Split(strOption, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "A"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = intStationQty - objKtbnStrc.strcSelection.intQuantity(9) - objKtbnStrc.strcSelection.intQuantity(10)
                    Case "F"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = intStationQty
                End Select
            Next

            'ねじ加算
            If bolScrewFlg Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           Right(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, 1)
                decOpAmount(UBound(decOpAmount)) = CDec(intScrewQty)
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Screw
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
