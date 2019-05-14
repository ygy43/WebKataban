'************************************************************************************
'*  ProgramID  ：KHPriceI8
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/27   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：直動形３・４方弁　ＦＳ＊／ＦＤ＊
'*
'************************************************************************************
Module KHPriceI8

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "FS1", "FS2", "FS3", "FS4", "FS5", _
                     "FD2", "FD3", "FD4", "FD5", "FDC3", _
                     "FDC4", "FDO3", "FDO4"

                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                        '基本価格キー
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                        'オプション加算価格キー
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case ""
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "FS" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim
                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case "P", "C"
                                            Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 1)
                                                Case "S"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                Case "D"
                                                    decOpAmount(UBound(decOpAmount)) = 2
                                            End Select
                                        Case Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                            End Select
                        Next
                    Else
                        '基本価格キー
                        If Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "C" Or _
                           Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1) = "O" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        'オプション加算価格キー
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case ""
                                Case Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "FS" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim
                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case "P", "C"
                                            Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 1)
                                                Case "S"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                Case "D"
                                                    decOpAmount(UBound(decOpAmount)) = 2
                                            End Select
                                        Case Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                            End Select
                        Next
                    End If
                Case "MFS2", "MFS3", "MFD2", "MFD3"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 3) & "-4"
                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(1).Trim

                    '固定加算(1)価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "KOTEIGAKU-1"
                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(1).Trim

                    '固定加算(2)価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "KOTEIGAKU-2"
                    decOpAmount(UBound(decOpAmount)) = 1
                    
                    'オプション加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "FS" & CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P", "C"
                                        Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 1)
                                            Case "S"
                                                decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(1).Trim)
                                            Case "D"
                                                decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(1).Trim) * 2
                                        End Select
                                    Case Else
                                        decOpAmount(UBound(decOpAmount)) = CDec(objKtbnStrc.strcSelection.strOpSymbol(1).Trim)
                                End Select
                        End Select
                    Next
                Case "FS2E", "FS3E", "FD2E", "FD3E"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "AFS1", "AFS2", "AFS3", "AFD2", "AFD3", "AFDC3", "AFDO3"
                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
