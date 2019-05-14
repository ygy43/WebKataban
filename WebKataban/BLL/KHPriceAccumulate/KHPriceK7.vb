'************************************************************************************
'*  ProgramID  ：KHPriceK7
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/08   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：Ｆ・Ｒ・Ｌ　ユニットモジュラータイプ
'*             ：Ｃ１０００／Ｃ１０１０／Ｃ１０２０／Ｃ１０３０／Ｃ１０４０／Ｃ１０５０／Ｃ１０６０
'*             ：Ｃ２５００／Ｃ２５２０／Ｃ２５３０／Ｃ２５５０
'*             ：Ｃ６５００／Ｃ６０２０／Ｃ６０３０／Ｃ６０５０／Ｃ６０６０／Ｃ６０７０
'*
'************************************************************************************
Module KHPriceK7

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

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

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
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "C1000", "C2500", "C6500"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    Select Case True
                        Case bolOptionF = True
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "F"

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
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "F1"

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
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "FF"

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
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "FF1"

                            If bolOptionT = True Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "T"
                            End If

                            If bolOptionY = True Then
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "Y"
                            End If

                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            If bolOptionT = True Or bolOptionY = True Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim

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

                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            'オプション加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(2), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   strOpArray(intLoopCnt).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
            Next

            '組付けアタッチメント加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen

                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                For intLoopCnt = 0 To strOpArray.Length - 1
                    Select Case strOpArray(intLoopCnt).Trim
                        Case ""
                        Case Else
                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & _
                                                                       strOpArray(intLoopCnt).Trim
                    End Select
                Next
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'アタッチメント加算価格キー
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case Else
                        Select Case True
                            Case Mid(strOpArray(intLoopCnt).Trim, 3, 1) = "N" Or Mid(strOpArray(intLoopCnt).Trim, 3, 1) = "G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "***" & _
                                                                           Left(strOpArray(intLoopCnt).Trim, 2)
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Mid(strOpArray(intLoopCnt).Trim, 4, 1) = "N" Or Mid(strOpArray(intLoopCnt).Trim, 4, 1) = "G"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "***" & _
                                                                           Left(strOpArray(intLoopCnt).Trim, 3)
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "***" & _
                                                                           strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
