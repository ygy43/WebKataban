'************************************************************************************
'*  ProgramID  ：KHPriceG6
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/01/25   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：薬液用エアオペレイトバルブ
'*             ：ＡＭＤ３＊２
'*             ：ＡＭＤ４＊２
'*             ：ＡＭＤ５＊２
'*             ：ＡＭＤ０＊２
'*
'************************************************************************************
Module KHPriceG6

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "AMD0"
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "Y" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            Case "0", "6"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "0" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "1"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "1" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "2", "7"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "2" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "3"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "3" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "8"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "8" & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Else
                        'RM1310067 2013/10/23
                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "" Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "0", "6"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "4-0"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "1"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "4-1"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "2", "7"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "4-2"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "3"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "4-3"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "8"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & "4-8"
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        Else
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "0", "6"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "0"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "1"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "1"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "2", "7"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "2"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "3"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "3"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case "8"
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & "8"
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        End If
                    End If
                Case "AMD3"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "1"
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "*" & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                "8" & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "Y" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                        strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                        objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            End If
                        Case "2"
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "*" & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                "10" & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "Y" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                        strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                        objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            End If
                        Case Else
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "*" & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen
                            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "Y" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                        strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            End If
                    End Select
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "AMD4"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "*" & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                "16" & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "Y" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                        strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                        objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                            End If
                        Case Else
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "*" & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen
                            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "Y" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                        strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & _
                                                                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                            End If
                    End Select
                    decOpAmount(UBound(decOpAmount)) = 1
                Case "AMD5"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case ""
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "*" & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                "20" & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case Else
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "*" & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                    End Select
                    decOpAmount(UBound(decOpAmount)) = 1
                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                                                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "*" & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(5).Trim

                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '流体加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "P" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "*" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                decOpAmount(UBound(decOpAmount)) = 1
            End If

            'AMD**2シリーズR,X追加 2008/5/2
            '↓RM1310067 2013/10/23
            If (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "AMD0" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "1") Or _
               (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "AMD3" And (objKtbnStrc.strcSelection.strKeyKataban.Trim = "1" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "2")) Or _
               (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "AMD4" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "") Or _
               (Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "AMD5" And objKtbnStrc.strcSelection.strKeyKataban.Trim = "") Then
                If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "R" Then

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                    Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                    "*" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                    "-" & objKtbnStrc.strcSelection.strOpSymbol(8).Trim

                    decOpAmount(UBound(decOpAmount)) = 1

                ElseIf objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "X" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <> 0 Then
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                        Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                        "*" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                        "-" & objKtbnStrc.strcSelection.strOpSymbol(7).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                        Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                        "*" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                        "- " & objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                    End If

                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "R,X" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                    Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                    "*" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & "-" & "R"

                    decOpAmount(UBound(decOpAmount)) = 1

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) <> 0 Then
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                        Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                        "*" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                        "-" & objKtbnStrc.strcSelection.strOpSymbol(7).Trim & "X"
                    Else
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                        Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & _
                        "*" & objKtbnStrc.strcSelection.strOpSymbol(2).Trim & _
                        "- " & "X"
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If
        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
