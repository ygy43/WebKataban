'************************************************************************************
'*  ProgramID  ：KHPriceN4
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/03/12   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ブロックマニホールドベースのみ　ＭＷ４ＧＢ４／ＭＷ４ＧＺ４
'*
'*  更新履歴   ：                       更新日：2007/10/04   更新者：NII A.Takahashi
'*               ・種類オプションに「CU」「CD」追加のため修正
'*               ・接続口径でねじ付を選択した場合、エンドブロックの両サイドにねじが付くため、
'*               　価格加算するよう修正
'*                                      更新日：2008/07/20      更新者：T.Sato
'*  ・受付No：RM0805028　種類オプションに「EU」「ED」追加のため修正
'*
'************************************************************************************
Module KHPriceN4

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim strOptionGN As String = ""
        Dim intScrewQty As Integer = 0

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            '基本価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "R1" Then
                Select Case Right(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1)
                    Case "G", "N"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        'Del by Zxjike 2013/10/03
                        'strOpRefKataban(UBound(strOpRefKataban)) = "N" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & "-V-" & _
                        '                                                 Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 2)
                        'Add by Zxjike 2013/10/03
                        strOpRefKataban(UBound(strOpRefKataban)) = "N" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & "-V-" & _
                                                                         Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 2) & _
                                                                         Right(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1)
                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)

                        'ねじ加算数量
                        'intScrewQty = intScrewQty + CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) Del by Zxjike 2013/10/03
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "N" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & "-V-" & _
                                                                         objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                End Select
            Else
                Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                    Case "1", "5"
                        Select Case Right(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1)
                            Case "G", "N"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'Del by Zxjike 2013/10/03
                                'strOpRefKataban(UBound(strOpRefKataban)) = "N" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & "-V1-" & _
                                '                                                 Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 2)
                                strOpRefKataban(UBound(strOpRefKataban)) = "N" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & "-V1-" & _
                                                 Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 2) & _
                                                 Right(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1)

                                decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)

                                'ねじ加算数量
                                'Del by Zxjike 2013/10/03
                                'intScrewQty = intScrewQty + CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) 
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "N" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & "-V1-" & _
                                                                                 objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                        End Select
                    Case Else
                        Select Case Right(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1)
                            Case "G", "N"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                'Del by Zxjike 2013/10/03
                                'strOpRefKataban(UBound(strOpRefKataban)) = "N" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & "-V2-" & _
                                '                                                 Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 2)
                                strOpRefKataban(UBound(strOpRefKataban)) = "N" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & "-V2-" & _
                                                 Left(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 2) & _
                                                 Right(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1)

                                decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)

                                'ねじ加算数量
                                'Del by Zxjike 2013/10/03
                                'intScrewQty = intScrewQty + CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "N" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & "-V2-" & _
                                                                                 objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                        End Select
                End Select
            End If

            '集中端子台・シリアル伝送子局
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "R1" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            '接続口径でG/Nを含んでいるか否か
            Select Case Right(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1)
                Case "G", "N"
                    strOptionGN = Right(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1)
            End Select

            'エンドブロック
            If InStr(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, "XU") <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-EXR"
                decOpAmount(UBound(decOpAmount)) = 1

                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-EL"
                decOpAmount(UBound(decOpAmount)) = 1

                If strOptionGN.Trim <> "" Then
                    intScrewQty = intScrewQty + 2
                End If
            ElseIf InStr(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, "XD") <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-EXL"
                decOpAmount(UBound(decOpAmount)) = 1

                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-ER"
                decOpAmount(UBound(decOpAmount)) = 1

                If strOptionGN.Trim <> "" Then
                    intScrewQty = intScrewQty + 2
                End If
            ElseIf InStr(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, "CU") <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-ECR"
                decOpAmount(UBound(decOpAmount)) = 1

                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-EL"
                decOpAmount(UBound(decOpAmount)) = 1

                If strOptionGN.Trim <> "" Then
                    intScrewQty = intScrewQty + 2
                End If
            ElseIf InStr(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, "CD") <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-ECL"
                decOpAmount(UBound(decOpAmount)) = 1

                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-ER"
                decOpAmount(UBound(decOpAmount)) = 1

                If strOptionGN.Trim <> "" Then
                    intScrewQty = intScrewQty + 2
                End If
            ElseIf InStr(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, "EU") <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-ENCR"
                decOpAmount(UBound(decOpAmount)) = 1

                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-EL"
                decOpAmount(UBound(decOpAmount)) = 1

                If strOptionGN.Trim <> "" Then
                    intScrewQty = intScrewQty + 2
                End If
            ElseIf InStr(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, "ED") <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-ENCL"
                decOpAmount(UBound(decOpAmount)) = 1

                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-ER"
                decOpAmount(UBound(decOpAmount)) = 1

                If strOptionGN.Trim <> "" Then
                    intScrewQty = intScrewQty + 2
                End If
            ElseIf InStr(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, "K") <> 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-ELK"
                decOpAmount(UBound(decOpAmount)) = 1

                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-ERK"
                decOpAmount(UBound(decOpAmount)) = 1

                If strOptionGN.Trim <> "" Then
                    intScrewQty = intScrewQty + 2
                End If
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-EL"
                decOpAmount(UBound(decOpAmount)) = 1

                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "NW4G4-ER"
                decOpAmount(UBound(decOpAmount)) = 1

                If strOptionGN.Trim <> "" Then
                    intScrewQty = intScrewQty + 2
                End If
            End If

            'オプション加算価格キー(その他)
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "F", "A"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5) & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                End Select
            Next

            'ねじ加算
            If intScrewQty > 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                           Right(objKtbnStrc.strcSelection.strOpSymbol(2).Trim, 1)
                decOpAmount(UBound(decOpAmount)) = CDec(intScrewQty)
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.Screw
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
