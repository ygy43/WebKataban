'************************************************************************************
'*  ProgramID  ：KHPriceR1
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2010/06/24   作成者：T.Fujiwara
'*                                      更新日：             更新者：
'*
'*  概要       ：小形直動式３方弁シリーズ  
'*       　　　　　　　  ３ＱＲＢ１　　　　　                                  
'*       　　　　　　　  Ｍ３ＱＲＡ１　　　　　                                
'*       　　　　　　　  Ｍ３ＱＲＢ１　
'*
'************************************************************************************
Module KHPriceR1

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim intStroke As Integer = 0
        Dim intLoopCnt As Integer
        Dim strOpArray() As String

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "3QRA1", "3QRB1"
                    '仕様なし(単体)
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '電線接続加算価格キー
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5)
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                    '流量サイズ加算価格キー
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(6)
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                Case "MV3QRA1", "MV3QRB1"
                    '仕様有(マニフォールド)
                    'サブプレート
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "SP" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(6)
                    decOpAmount(UBound(decOpAmount)) = 1

                    Dim intQuantity As Integer = 0



                    For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                        If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then


                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4) = "+ｾﾝｻ" OrElse Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 7) = "+Senser" Then
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 7)
                            Else
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                            End If
                            decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)



                            '2010/12/10 MOD RM1012055(1月VerUP:3QRシリーズ) START--->
                            If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 6, 1) <> "M" Then
                                'If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 7, 1) <> "M" Then
                                '2010/12/10 MOD RM1012055(1月VerUP:3QRシリーズ) <---END
                                If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 5, 1) = "1" Then
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                End If
                            End If
                            Dim strKt As String = Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 5)
                            Dim strKt2 As String = Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 6, 1)
                            '電線接続加算価格キー
                            If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                                If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 2) <> "MP" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)


                                    strOpRefKataban(UBound(strOpRefKataban)) = strKt.Trim & CdCst.Sign.Hypen & _
                                                                                 strKt2.Trim & CdCst.Sign.Hypen & _
                                                                                objKtbnStrc.strcSelection.strOpSymbol(4)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                End If
                            End If
                            '流量サイズ加算価格キー
                            If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                                If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 2) <> "MP" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strKt.Trim & CdCst.Sign.Hypen & _
                                                                                objKtbnStrc.strcSelection.strOpSymbol(5)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                            End If

                            '圧力センサ加算価格キー
                            If Len(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) <> 0 Then
                                If Right(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 2) <> "MP" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = strKt.Trim & CdCst.Sign.Hypen & _
                                                                                objKtbnStrc.strcSelection.strOpSymbol(8)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                            End If
                        End If

                    Next

                    Dim strWk As String = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3, 5).Trim

                    ''マニホールド以外基本加算価格
                    'If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "8" Then
                    '    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '    strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & objKtbnStrc.strcSelection.strOpSymbol(1) & "9"
                    '    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    'End If

                    '接続口径加算価格キー
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) <> 0 Then

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(3)
                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(6).Trim

                    End If

                    ''電線接続加算価格キー
                    'If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                    '    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then

                    '    Else
                    '        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    '        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    '        strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                    '                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                    '                                                   objKtbnStrc.strcSelection.strOpSymbol(4)
                    '        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    '    End If
                    'End If

                Case "3QB1"
                    '仕様なし(単体)
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '電線接続加算価格キー
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4)
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                    '流量サイズ加算価格キー
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5)
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                    '圧力仕様加算価格キー
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(6)
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                Case "3QE1"
                    '仕様なし(単体)
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(2) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '手動装置加算価格キー
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(4)
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If


                    '電線接続加算価格キー
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                   CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(5)
                        decOpAmount(UBound(decOpAmount)) = 1

                    End If

                    'オプション加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                          Select strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    Next


                Case "M3QB1"
                    '仕様有(マニフォールド)
                    'サブプレート
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "SP" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7)
                    decOpAmount(UBound(decOpAmount)) = 1

                    Dim intQuantity As Integer = 0

                    For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                        If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                                If intLoopCnt = 2 Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                End If
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End If

                            If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 6, 1) <> "M" Then
                                If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 1) = "1" Then
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                End If
                            End If
                        End If
                    Next

                    Dim strWk As String = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5).Trim

                    '電線接続加算価格キー
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                        "1" & CdCst.Sign.Hypen & _
                                                                        objKtbnStrc.strcSelection.strOpSymbol(4)
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4)
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        End If
                    End If

                    '流量サイズ加算価格キー
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(5)
                        decOpAmount(UBound(decOpAmount)) = intQuantity

                    End If

                    '圧力仕様加算価格キー
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <> 0 Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(6)
                        decOpAmount(UBound(decOpAmount)) = intQuantity
                    End If

                Case "M3QE1", "M3QZ1"
                    '仕様有(マニフォールド)
                    'サブプレート
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "SP" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7)
                    decOpAmount(UBound(decOpAmount)) = 1

                    Dim intQuantity As Integer = 0

                    For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                        If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    If intLoopCnt = 2 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                End If
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End If

                            If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 6, 1) <> "M" Then
                                If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 4, 1) = "1" Then
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                End If
                            End If
                        End If
                    Next

                    Dim strWk As String = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5).Trim

                    '手動装置加算価格キー
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) <> 0 Then

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(4)
                        decOpAmount(UBound(decOpAmount)) = intQuantity
                    End If

                    '電線接続加算価格キー
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                        "1" & CdCst.Sign.Hypen & _
                                                                        objKtbnStrc.strcSelection.strOpSymbol(5)
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5)
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        End If
                    End If

                    'オプション加算価格キー
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(6), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                            strOpArray(intLoopCnt).Trim
                                decOpAmount(UBound(decOpAmount)) = intQuantity
                        End Select
                    Next

                Case Else
                    '仕様有(マニフォールド)
                    'サブプレート
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "SP" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(7)
                    decOpAmount(UBound(decOpAmount)) = 1

                    Dim intQuantity As Integer = 0

                    For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                        If objKtbnStrc.strcSelection.intQuantity(intLoopCnt) <> 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim <> "" And _
                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "H" And _
                                   objKtbnStrc.strcSelection.strOpSymbol(9).Trim <> "ST" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    If intLoopCnt = 2 Then
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    Else
                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                        decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    End If
                                End If
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt)
                                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End If

                            '2010/12/10 MOD RM1012055(1月VerUP:3QRシリーズ) START--->
                            If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 6, 1) <> "M" Then
                                'If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 7, 1) <> "M" Then
                                '2010/12/10 MOD RM1012055(1月VerUP:3QRシリーズ) <---END
                                If Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt), 5, 1) = "1" Then
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                Else
                                    intQuantity = intQuantity + objKtbnStrc.strcSelection.intQuantity(intLoopCnt) * 2
                                End If
                            End If
                        End If
                    Next

                    Dim strWk As String = Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2, 5).Trim

                    '電線接続加算価格キー
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) <> 0 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "8" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                        "1" & CdCst.Sign.Hypen & _
                                                                        objKtbnStrc.strcSelection.strOpSymbol(5)
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5)
                            decOpAmount(UBound(decOpAmount)) = intQuantity
                        End If
                    End If

                    '流量サイズ加算価格キー
                    If Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) <> 0 Then

                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = strWk.Trim & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(6)
                        decOpAmount(UBound(decOpAmount)) = intQuantity

                    End If
            End Select

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module

