'************************************************************************************
'*  ProgramID  ：KHPriceL1
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/30   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：小形直動精密レギュレータ
'*             ：ＲＪＢ５００
'*             ：ＭＮＲＪＢ５００
'*             ：ＮＲＪＢ５００
'*
'************************************************************************************
Module KHPriceL1

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String = Nothing
        Dim intLoopCnt As Integer
        Dim bolOptionL As Boolean
        Dim bolOptionT As Boolean
        Dim bolFirst As Boolean

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "RJB500"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "1"
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'オプション加算価格キー
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                            bolOptionL = False
                            bolOptionT = False
                            bolFirst = True
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "L"
                                        bolOptionL = True
                                    Case "T"
                                        bolOptionT = True
                                End Select
                            Next
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case "L", "T"
                                        If bolOptionL = True And bolOptionT = True Then
                                            If bolFirst = True Then
                                                bolFirst = False
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "LT"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                       strOpArray(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                        Case "2"
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'オプション加算価格キー
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                            bolOptionL = False
                            bolOptionT = False
                            bolFirst = True
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "L"
                                        bolOptionL = True
                                    Case "T"
                                        bolOptionT = True
                                End Select
                            Next
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case "L", "T"
                                        If bolOptionL = True And bolOptionT = True Then
                                            If bolFirst = True Then
                                                bolFirst = False
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "LT"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                       strOpArray(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                    End Select
                Case "MNRJB500"
                    '基本価格キー
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "A"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                       Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) & _
                                                                       Right(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1)
                            decOpAmount(UBound(decOpAmount)) = CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim)
                        Case "B"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                       Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1) & _
                                                                       Right(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1)
                            decOpAmount(UBound(decOpAmount)) = CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim)
                    End Select

                    'エンドブロック(右側)加算価格キー
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        Case "D"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & CdCst.Sign.Hypen & "NE" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & CdCst.Sign.Hypen & "NE"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'エンドブロック(左側)加算価格キー
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        Case "D"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & CdCst.Sign.Hypen & "NE" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim & "L"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & CdCst.Sign.Hypen & "NEL"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'ＤＩＮレール加算価格キー
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                        Case "D"
                        Case Else
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                Case "A"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "1"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA125"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "2"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA150"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "3"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA175"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "4"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA212.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "5"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA237.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "6"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA262.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "7"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA287.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "8"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA325"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "9"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA350"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "10"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA375"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                Case "B"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "1"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA100"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "2"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA137.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "3"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA162.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "4"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA187.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "5"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA212.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "6"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA250"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "7"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA275"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "8"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA300"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "9"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA325"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case "10"
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 2) & _
                                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5, 4) & _
                                                                                       CdCst.Sign.Hypen & "BAA362.5"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                            End Select
                    End Select

                    '集中給気ブロック加算価格キー
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "A"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & CdCst.Sign.Hypen & "NP" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                       Left(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 2)
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

                    'オプション加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
                    bolOptionL = False
                    bolOptionT = False
                    bolFirst = True
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "L"
                                bolOptionL = True
                            Case "T"
                                bolOptionT = True
                        End Select
                    Next
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "L", "T"
                                If bolOptionL = True And bolOptionT = True Then
                                    If bolFirst = True Then
                                        bolFirst = False
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "LT"
                                        decOpAmount(UBound(decOpAmount)) = CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim)
                                    End If
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                               strOpArray(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim)
                                End If
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim)
                        End Select
                    Next
                Case "NRJB500"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "1"
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'オプション加算価格キー
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(5), CdCst.Sign.Delimiter.Comma)
                            bolOptionL = False
                            bolOptionT = False
                            bolFirst = True
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "L"
                                        bolOptionL = True
                                    Case "T"
                                        bolOptionT = True
                                End Select
                            Next
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case "L", "T"
                                        If bolOptionL = True And bolOptionT = True Then
                                            If bolFirst = True Then
                                                bolFirst = False
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "LT"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            End If
                                        Else
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                                       strOpArray(intLoopCnt).Trim
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        End If
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                        Case "2"
                            '基本価格キー
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            decOpAmount(UBound(decOpAmount)) = 1

                            'オプション加算価格キー
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case ""
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                                   strOpArray(intLoopCnt).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Next
                    End Select
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
