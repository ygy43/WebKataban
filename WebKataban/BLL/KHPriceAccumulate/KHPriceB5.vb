'************************************************************************************
'*  ProgramID  ：KHPriceB5
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/12   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：省配線ブロックマニホールド　ＭＷ３ＧＡ２／ＭＷ４ＧＡ２／ＭＷ４ＧＢ２／ＭＷ４ＧＺ２
'*
'************************************************************************************
Module KHPriceB5

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim intValveQty As Integer = 0
        Dim intValveQty3P As Integer = 0
        Dim intValveQty4P As Integer = 0
        Dim intValveQty1SWD As Integer = 0
        Dim intValveQty3SWD As Integer = 0
        Dim intValveQty5SWD As Integer = 0
        Dim intValveQty5SWD2 As Integer = 0

        Dim bolSupplySA As Boolean = False
        Dim bolSupplyS As Boolean = False
        
        Dim intValveQtyMP As Integer = 0
        Dim intStationQty As Integer = 0
        Dim bolOptionH As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            'バルブブロック連数
            intStationQty = CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
            
            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                   objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                    '2012/10/26 初期化する
                    bolSupplySA = False
                    bolSupplyS = False
                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                        Case CdCst.Manifold.InspReportJp.SelectValue, _
                             CdCst.Manifold.InspReportEn.SelectValue, _
                            CdCst.Manifold.InspReportJp.English, CdCst.Manifold.InspReportEn.English, _
                            CdCst.Manifold.InspReportJp.Japanese, CdCst.Manifold.InspReportEn.Japanese
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

                                    'FP1シリーズ追加対応  2017/02/15 追加
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "O", "X"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = "NW4G2-" & Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 8, Len(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim)) & "-FP1"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End Select

                                Case 3
                                    '電装ブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'FP1シリーズ追加対応  2017/02/15 追加
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "O", "X"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End Select
                                Case 4 To 11
                                    '電磁弁付バルブブロック＆MP付バルブブロック
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-MP") = 0 Then
                                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "R1" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            'FP1シリーズ追加対応  2017/02/15 追加
                                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                Case "O", "X"
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                    'バルブブロックキーの作成を追加  2017/03/07 追加
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-V-FP1"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End Select

                                            'RM1805036_二次電池加算価格対応
                                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                Case "P", "Y"
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-P40"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End Select

                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            'FP1シリーズ追加対応  2017/02/15 追加
                                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                Case "O", "X"
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                    'バルブブロックキーの作成を追加  2017/03/07 追加
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-V-FP1"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            End Select

                                            'RM1805036_二次電池加算価格対応
                                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                Case "P", "Y"
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-P40"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End Select

                                        End If

                                        '電磁弁数(バルブ数)をカウントする①
                                        intValveQty = intValveQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                        '電磁弁数(バルブ数)をカウントする②
                                        Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 3, 1)
                                            Case "3"
                                                intValveQty3P = intValveQty3P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Case "4"
                                                intValveQty4P = intValveQty4P + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End Select

                                        '電磁弁数(バルブ数)をカウントする③
                                        Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 7, 1)
                                            Case "1"
                                                intValveQty1SWD = intValveQty1SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Case "2", "3", "4", "5"
                                                intValveQty5SWD2 = intValveQty5SWD2 + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Case "3"
                                                intValveQty3SWD = intValveQty3SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            Case "5"
                                                intValveQty5SWD = intValveQty5SWD + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End Select

                                        '電磁弁数(バルブ数)をカウントする③
                                        intValveQtyMP = intValveQtyMP + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        'MP付バルブブロックの時
                                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "R1" Then
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            'FP1シリーズ追加対応  2017/02/15 追加
                                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                Case "O", "X"
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                    'バルブブロックキーの作成を追加  2017/03/07 追加
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-V-FP1"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            End Select

                                            'RM1805036_二次電池加算価格対応
                                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                Case "P", "Y"
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-P40"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End Select

                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            'FP1シリーズ追加対応  2017/02/15 追加
                                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                Case "O", "X"
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                    'バルブブロックキーの作成を追加  2017/03/07 追加
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-V-FP1"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                            End Select

                                            'RM1805036_二次電池加算価格対応
                                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                Case "P", "Y"
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-P40"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End Select

                                        End If

                                        'MP数(バルブ数)をカウントする③
                                        intValveQtyMP = intValveQtyMP + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                    'RM1002027 2010/02/16 Y.Miura 仕様画面のオプション加算不具合対応
                                    'Case 12 To 13
                                Case 12 To 15
                                    'スペーサ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "O", "X"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End Select

                                    'Case 14 To 15
                                Case 16 To 17
                                    '給排気ブロック
                                    '仕切タイプ給排気ブロック選択時
                                    If InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-SA") <> 0 Or _
                                       InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-S") <> 0 Then
                                        '2012/10/26 数量判断を追加する
                                        If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
                                            Select Case True
                                                Case InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-SA") <> 0
                                                    bolSupplySA = True

                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-SA") - 1)
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                    'FP1シリーズ追加対応  2017/02/15 追加
                                                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-8") Or
                                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-8X") Or
                                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-10") Or
                                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-10X") Then

                                                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                            Case "O", "X"
                                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-SA") - 1) & "-FP1"
                                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        End Select

                                                    End If

                                                    'RM1805036_二次電池加算価格対応
                                                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-8") Or
                                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-8X") Or
                                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-10") Or
                                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-10X") Then

                                                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                            Case "P", "Y"
                                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-SA") - 1) & "-P40"
                                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        End Select

                                                    End If

                                                Case InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-S") <> 0
                                                    bolSupplyS = True

                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-S") - 1)
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                                    'FP1シリーズ追加対応  2017/02/15 追加
                                                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-8") Or
                                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-8X") Or
                                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-10") Or
                                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-10X") Then

                                                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                            Case "O", "X"
                                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-S") - 1) & "-FP1"
                                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        End Select

                                                    End If

                                                    'RM1805036_二次電池加算価格対応
                                                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-8") Or
                                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-8X") Or
                                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-10") Or
                                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-10X") Then

                                                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                            Case "P", "Y"
                                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 1, InStr(1, objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, "-S") - 1) & "-P40"
                                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                        End Select

                                                    End If

                                            End Select
                                        End If
                                    Else
                                        '通常
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                        'FP1シリーズ追加対応  2017/02/15 追加
                                        If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-8") Or
                                           objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-8X") Or
                                           objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-10") Or
                                           objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-10X") Then

                                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                Case "O", "X"
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End Select

                                        End If

                                        'RM1805036_二次電池加算価格対応
                                        If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-8") Or
                                           objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-8X") Or
                                           objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-10") Or
                                           objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-Q-10X") Then

                                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                                Case "P", "Y"
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-P40"
                                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End Select

                                        End If

                                    End If
                                    'Case 16 To 17
                                Case 18 To 19
                                    '仕切りブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'FP1シリーズ追加対応  2017/02/15 追加
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "O", "X"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End Select

                                    'Case 18 To 19
                                Case 20 To 21
                                    'エンドブロック
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'FP1シリーズ追加対応  2017/02/15 追加
                                    If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-ER") Or
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-EXR") Or
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-EL") Or
                                       objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Contains("-EXL") Then
                                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                            Case "O", "X"
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim & "-FP1"
                                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                        End Select
                                    End If

                                    'Case 20 To 22
                                Case 22 To 24
                                    'ブランクプラグ、サイレンサ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-" & objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    'Case 23
                                Case 25
                                    '防水プラグ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    'Case 24 To 25  'RM1002027 2010/02/16 Y.Miura 仕様画面のオプション加算不具合対応
                                Case 26 To 27
                                    '検査成績書＆ケーブル
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    'Case 26
                                Case 28
                                    'ケーブルクランプ
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 11)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    'Case 27
                                Case 29
                                    'タグ銘板
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-TAG"
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End Select

                            '仕切りタイプ給排気ブロック("-SA","-S")選択加算
                            If bolSupplySA = True Or _
                               bolSupplyS = True Then
                                Select Case True
                                    Case bolSupplySA = True
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-SA"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Case bolSupplyS = True
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-S"
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End Select
                            End If
                    End Select
                End If
            Next

            'DINレール加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "D" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "N4G-BAA"
                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.decDinRailLength
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail
            End If

            'DINレールマウントタイプ加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "D" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "W4G2" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                decOpAmount(UBound(decOpAmount)) = 1

                'FP1シリーズ追加対応  2017/02/15 追加
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "O", "X"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "W4G2" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(7).Trim & "-FP1"
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select

            End If

            ''食品製造工程向けオプション追加に伴う処理の追加  2017/02/15 追加
            'If objKtbnStrc.strcSelection.strKeyKataban = "O" Or objKtbnStrc.strcSelection.strKeyKataban = "X" Then

            '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
            '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
            '    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
            '    strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-FP1"
            '    decOpAmount(UBound(decOpAmount)) = 1

            'End If

            '電圧が４の場合のみ価格キーを作成するよう変更  2017/03/03 追加
            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "4" Then
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "O", "X"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-DC-FP1"
                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.decDinRailLength
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail
                End Select
            End If

            '電圧が１の場合のみ価格キーを作成するよう変更  2017/03/03 追加
            If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "1" Then
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "O", "X"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-AC-FP1"
                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.decDinRailLength
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail
                End Select
            End If

            'オプション　加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "A"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = intValveQty
                    Case "F"
                        Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                            Case "MW3GA2", "MW4GB2", "MW4GZ2"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim
                                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                    Case "MW3GA2", "MW3GB2", "MW3GZ2"
                                        decOpAmount(UBound(decOpAmount)) = intValveQty
                                    Case "MW4GB2", "MW4GZ2"
                                        decOpAmount(UBound(decOpAmount)) = intValveQtyMP
                                End Select
                            Case "MW4GA2"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "W3GA2-" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intValveQty3P

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "W4GA2-" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intValveQty4P
                        End Select
                    Case "M7"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-" & strOpArray(intLoopCnt).Trim & "-S"
                        decOpAmount(UBound(decOpAmount)) = intValveQty1SWD

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-" & strOpArray(intLoopCnt).Trim & "-D"
                        decOpAmount(UBound(decOpAmount)) = intValveQty5SWD2

                    Case "M"
                        'オプションＭを追加
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "O", "X"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "W4G2-" & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intValveQty1SWD
                        End Select

                    Case "H"
                        bolOptionH = True

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & "-H"
                        decOpAmount(UBound(decOpAmount)) = intValveQty3SWD + intValveQty5SWD
                End Select
            Next
            If bolOptionH = False Then
                '電磁弁バルブブロックから、排気誤作動防止弁価格を減算するためのキーを生成
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & CdCst.Sign.Hypen & "H"
                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    Case "MW3GA2", "MW3GB2", "MW3GZ2"
                        decOpAmount(UBound(decOpAmount)) = intValveQty - intValveQty3SWD - intValveQty5SWD
                    Case "MW4GA2", "MW4GB2", "MW4GZ2"
                        decOpAmount(UBound(decOpAmount)) = intValveQty
                End Select
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
