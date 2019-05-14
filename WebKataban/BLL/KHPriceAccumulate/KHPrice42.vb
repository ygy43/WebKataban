'************************************************************************************
'*  ProgramID  ：KHPrice42
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/23   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：ＲＧ＊＊／ＰＣ＊Ｓ
'*
'************************************************************************************
Module KHPrice42

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '価格キー設定
            Select Case True
                Case objKtbnStrc.strcSelection.strKeyKataban.Trim = "W"
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 1) >= "0" And Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 1) <= "9" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                                   Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & CdCst.Sign.Hypen & "W" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "FC" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(11).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(12).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                                   Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & CdCst.Sign.Hypen & "W" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "AL" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(11).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(12).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case objKtbnStrc.strcSelection.strKeyKataban.Trim = "E"
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 1) >= "0" And Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 1) <= "9" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                                   Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & CdCst.Sign.Hypen & "E" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "FC" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(11).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(12).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                                   Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & CdCst.Sign.Hypen & "E" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "AL" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(11).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(12).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case objKtbnStrc.strcSelection.strKeyKataban.Trim = "G"
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 1) >= "0" And Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 1) <= "9" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                                   Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & CdCst.Sign.Hypen & "G" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "FC" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(11).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(12).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                                   Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & CdCst.Sign.Hypen & "G" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "AL" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(11).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(12).Trim & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(15).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case Else
                    If Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 1) >= "0" And Left(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 1) <= "9" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                                   Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "FC"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                                   Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "AL"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

            '入出力仕様加算価格キー
            For intLoopCnt = 8 To 10
                If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim <> "" And _
                   objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim <> "N" And _
                   (intLoopCnt <> 9 Or objKtbnStrc.strcSelection.strOpSymbol(8).Trim <> "K" Or objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "K") Then
                    Select Case True
                        Case intLoopCnt = 10
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                                       Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "O" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim = "H" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                                           Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "I" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                                           Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "I" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
                End If
            Next

            'オプション加算価格キー
            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                Case "W", "E"
                    intLoopCnt = 15
                Case "G", "H"
                    intLoopCnt = 16
                Case Else
                    intLoopCnt = 11
            End Select

            If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & "*" & _
                                                           Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim

                Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                    Case "F", "A", "S", "B"
                        Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1)
                            Case "S", "L"
                                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                    Case "RGCS", "RGIL", "RGIS", "RGOL", "RGOS", "PCIS", "PCOS"
                                        If objKtbnStrc.strcSelection.strOpSymbol(12).Trim <> "" Then
                                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "TSF" & _
                                                                                       objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                                        Else
                                            strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "TSF" & CdCst.Sign.Hypen & "NO"
                                        End If
                                    Case Else
                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "TSF" & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                                End Select
                            Case "T"
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "TST" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                        End Select
                    Case "X", "C", "Y", "D"
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & "TGX" & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                End Select

                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                decOpAmount(UBound(decOpAmount)) = 1
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
