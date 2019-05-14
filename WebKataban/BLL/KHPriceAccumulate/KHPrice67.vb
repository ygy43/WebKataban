'************************************************************************************
'*  ProgramID  ：KHPrice67
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/22   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：セレックスロータリー
'*             ：ＲＶ３Ｓ／ＤＡ　小形・角度可変形
'*             ：ＲＶ３Ｓ／ＤＨ　大形・低油圧形
'*
'************************************************************************************
Module KHPrice67

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)


        Dim strOpArray() As String
        Dim intLoopCnt As Integer
        Dim bolOptionC As Boolean = False

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "RV3SA", "RV3DA"
                    '基本価格キー
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "0"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "HOPEANGLE" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select


                    'スイッチ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "RV3S" & CdCst.Sign.Hypen & "FR" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'オプション加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
                Case "RV3SH", "RV3DH"
                    '選択オプション分解＆ショックキラー選択判定
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "C"
                                bolOptionC = True
                        End Select
                    Next

                    '基本価格キー
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    decOpAmount(UBound(decOpAmount)) = 1

                    'スイッチ加算価格キー
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "" Then
                        Select Case bolOptionC
                            Case True
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "D"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "C" & CdCst.Sign.Hypen & _
                                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                                           objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & "C" & CdCst.Sign.Hypen & _
                                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                            Case Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                    Case "D"
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                                           objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                    Case Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                        End Select

                        'リード線長さ加算価格キー
                        If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = "RVU" & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                        End If
                    End If

                    'オプション加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "FA", "LS"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & CdCst.Sign.Hypen & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "C"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVC" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                decOpAmount(UBound(decOpAmount)) = 1

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = "RVC" & objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                                   objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & "T"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next
            End Select

        Catch ex As Exception

            Throw ex

        Finally



        End Try

    End Sub

End Module
