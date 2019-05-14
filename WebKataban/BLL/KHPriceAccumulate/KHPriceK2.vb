'************************************************************************************
'*  ProgramID  ：KHPriceK2
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/08   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：オイルミストフィルタ　Ｍ６０００
'*
'************************************************************************************
Module KHPriceK2

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim bolOptionF1 As Boolean = False
        Dim bolOptionS As Boolean = False
        Dim bolOptionX As Boolean = False
        Dim strBoreSign As String

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

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
                    Case "F1"
                        bolOptionF1 = True
                    Case "S"
                        bolOptionS = True
                    Case "X"
                        bolOptionX = True
                End Select
            Next

            '基本価格キー
            Select Case bolOptionF1
                Case True
                    Select Case True
                        Case bolOptionS = True
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       strBoreSign & CdCst.Sign.Hypen & "F1S"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case bolOptionX = True
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       strBoreSign & CdCst.Sign.Hypen & "F1X"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       strBoreSign & CdCst.Sign.Hypen & "F1"
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case Else
                    Select Case True
                        Case bolOptionS = True
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       strBoreSign & CdCst.Sign.Hypen & "S"
                        Case bolOptionX = True
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       strBoreSign & CdCst.Sign.Hypen & "X"
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       strBoreSign
                    End Select

                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "P70", "P74"
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    End Select
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(2), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "Z", "M"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "C6", "M6"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "Q"
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "P70", "P74"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select
            Next

            'アタッチメント加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(5), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "A20", "A25", "A32", "A20N", "A25N", _
                         "A32N", "A20G", "A25G", "A32G", "B"
                        Select Case True
                            Case Mid(strOpArray(intLoopCnt).Trim, 3, 1) = "N" Or Mid(strOpArray(intLoopCnt).Trim, 3, 1) = "G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           Left(strOpArray(intLoopCnt).Trim, 2)
                            Case Mid(strOpArray(intLoopCnt).Trim, 4, 1) = "N" Or Mid(strOpArray(intLoopCnt).Trim, 4, 1) = "G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           Left(strOpArray(intLoopCnt).Trim, 3)
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           strOpArray(intLoopCnt).Trim
                        End Select

                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "P70", "P74"
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        End Select
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
