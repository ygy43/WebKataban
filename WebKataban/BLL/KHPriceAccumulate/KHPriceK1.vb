'************************************************************************************
'*  ProgramID  ：KHPriceK1
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/08   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：エアフィルタ　Ｆ６０００
'*
'************************************************************************************
Module KHPriceK1

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim bolOptionF As Boolean = False
        Dim bolOptionF1 As Boolean = False
        Dim bolOptionFF As Boolean = False
        Dim bolOptionFF1 As Boolean = False
        Dim bolOptionY As Boolean = False
        Dim bolOptionT As Boolean = False
        Dim bolOptionC6 As Boolean = False
        Dim bolOptionM6 As Boolean = False
        Dim bolOptionQ As Boolean = False
        Dim strBoreSign As String

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "P74" Then
                'オプション選択判定
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(2), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case "Y"
                            bolOptionY = True
                        Case "C6"
                            bolOptionC6 = True
                        Case "M6"
                            bolOptionM6 = True
                        Case "Q"
                            bolOptionQ = True
                    End Select
                Next

                '基本価格キー
                If bolOptionY = True Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "Y" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

                'C6オプション加算価格キー
                If bolOptionC6 = True Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "C6"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

                'M6オプション加算価格キー
                If bolOptionM6 = True Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "M6"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

                'Qオプション加算価格キー
                If bolOptionQ = True Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "Q"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

                'アタッチメント加算価格キー
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(5), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       strOpArray(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Next
            Else
                '接続口径サイズ判定
                Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    Case "20G", "25G"
                        strBoreSign = "G"
                    Case Else
                        strBoreSign = ""
                End Select

                'オプション選択判定
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(2), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case "F"
                            bolOptionF = True
                        Case "F1"
                            bolOptionF1 = True
                        Case "FF"
                            bolOptionFF = True
                        Case "FF1"
                            bolOptionFF1 = True
                        Case "Y"
                            bolOptionY = True
                        Case "T"
                            bolOptionT = True
                    End Select
                Next

                '基本価格キー
                Select Case True
                    Case bolOptionF = True
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                   "*00" & strBoreSign & CdCst.Sign.Hypen & "F"

                        If bolOptionT = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                        End If

                        If bolOptionY = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                        End If
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case bolOptionF1 = True
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                   "*00" & strBoreSign & CdCst.Sign.Hypen & "F1"

                        If bolOptionT = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                        End If

                        If bolOptionY = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                        End If
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case bolOptionFF = True
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                   "*00" & strBoreSign & CdCst.Sign.Hypen & "FF"

                        If bolOptionT = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                        End If

                        If bolOptionY = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                        End If
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case bolOptionFF1 = True
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                   "*00" & strBoreSign & CdCst.Sign.Hypen & "FF1"

                        If bolOptionT = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                        End If

                        If bolOptionY = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                        End If
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                   "*00" & strBoreSign

                        If bolOptionT = True Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "T"
                        End If

                        If bolOptionY = True Then
                            If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 6, 1) = "-" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                            Else
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "Y"
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "P70" Then
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        End If
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select

                'オプション加算価格キー
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(2), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                       "*00" & strOpArray(intLoopCnt).Trim

                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "P70" Then
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "C6", "M6", "Q"
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                End Select
                            End If
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Next

                'アタッチメント加算価格キー
                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(5), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else
                            Select Case True
                                Case Mid(strOpArray(intLoopCnt).Trim, 3, 1) = "N" Or Mid(strOpArray(intLoopCnt).Trim, 3, 1) = "G"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                               "*00" & Left(strOpArray(intLoopCnt).Trim, 2)
                                Case Mid(strOpArray(intLoopCnt).Trim, 4, 1) = "N" Or Mid(strOpArray(intLoopCnt).Trim, 4, 1) = "G"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                               "*00" & Left(strOpArray(intLoopCnt).Trim, 3)
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                               "*00" & strOpArray(intLoopCnt).Trim
                            End Select

                            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "P70" Then
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "A20", "A25", "A32", "A20N", "A25N", _
                                         "A32N", "A20G", "A25G", "A32G", "B"
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                End Select
                            End If
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Next
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
