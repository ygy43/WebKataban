'************************************************************************************
'*  ProgramID  ：KHPriceQ0
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2008/06/10   作成者：M.Kojima
'*
'*  概要       ：自動調圧式バランサシリンダ BBS-A,BBS-O/OBシリーズ
'*
'************************************************************************************
Module KHPriceQ0
    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)


        Dim intStroke As Integer = 0

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "BBS-A" Then
                '基本価格キーの設定
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                'RM1410046 BBSストロークキー変更
                If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 100 Then
                    '~100M
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        "100"
                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 100 And objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 200 Then
                    '101M~200M
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        "200"
                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 200 And objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 300 Then
                    '201M~300M
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        "300"
                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 300 And objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 400 Then
                    '301M~400M
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        "400"
                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 400 And objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 500 Then
                    '401M~500M
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        "500"
                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 500 And objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 600 Then
                    '501M~600M
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        "600"
                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 600 And objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 700 Then
                    '601M~700M
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        "700"
                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 700 And objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 800 Then
                    '701M~800M
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        "800"
                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 800 And objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 900 Then
                    '801M~900M
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        "900"
                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 900 And objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 1000 Then
                    '901M~1000M
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        "1000"
                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 1000 And objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 1100 Then
                    '1001M~1100M
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        "1100"
                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 1100 And objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 1200 Then
                    '1101M~1200M
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        "1200"
                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 1200 And objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 1300 Then
                    '1201M~1300M
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        "1300"
                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 1300 And objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 1400 Then
                    '1301M~1400M
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        "1400"
                    decOpAmount(UBound(decOpAmount)) = 1
                ElseIf objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 1400 And objKtbnStrc.strcSelection.strOpSymbol(7).Trim <= 1500 Then
                    '1401M~1500M
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                        CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        "1500"
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

                '支持加算
                If (objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "00") Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                    CdCst.Sign.Hypen & _
                    objKtbnStrc.strcSelection.strOpSymbol(4).Trim & CdCst.Sign.Hypen & _
                    objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                    objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

                '付属品加算
                If objKtbnStrc.strcSelection.strOpSymbol(8).Trim.Length <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            ElseIf objKtbnStrc.strcSelection.strSeriesKataban.Trim = "BBS-O" Or _
                    objKtbnStrc.strcSelection.strSeriesKataban.Trim = "BBS-OB" Then
                '基本価格キーの設定
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)

                'RM1410046  BBSストロークキー変更
                Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                    Case "BBS-O"
                        If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 100 Then
                            '~100M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "100"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 100 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 200 Then
                            '101M~200M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "200"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 200 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 300 Then
                            '201M~300M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "300"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 300 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 400 Then
                            '301M~400M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "400"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 400 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 500 Then
                            '401M~500M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "500"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 500 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 600 Then
                            '501M~600M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "600"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 600 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 700 Then
                            '601M~700M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "700"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 700 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 800 Then
                            '701M~800M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "800"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 800 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 900 Then
                            '801M~900M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "900"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 900 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 1000 Then
                            '901M~1000M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "1000"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 1000 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 1100 Then
                            '1001M~1100M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "1100"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 1100 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 1200 Then
                            '1101M~1200M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "1200"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 1200 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 1300 Then
                            '1201M~1300M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "1300"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 1300 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 1400 Then
                            '1301M~1400M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "1400"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 1400 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 1500 Then
                            '1401M~1500M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "1500"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    Case "BBS-OB"
                        If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 100 Then
                            '~100M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "100"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 100 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 200 Then
                            '101M~200M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "200"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 200 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 300 Then
                            '201M~300M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "300"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 300 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 400 Then
                            '301M~400M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "400"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 400 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 500 Then
                            '401M~500M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "500"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 500 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 600 Then
                            '501M~600M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "600"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 600 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 700 Then
                            '601M~700M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "700"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 700 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 800 Then
                            '701M~800M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "800"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 800 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 900 Then
                            '801M~900M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "900"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 900 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 1000 Then
                            '901M~1000M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "1000"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 1000 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 1100 Then
                            '1001M~1100M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "1100"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 1100 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 1200 Then
                            '1101M~1200M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "1200"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 1200 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 1300 Then
                            '1201M~1300M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "1300"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 1300 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 1400 Then
                            '1301M~1400M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "1400"
                            decOpAmount(UBound(decOpAmount)) = 1
                        ElseIf objKtbnStrc.strcSelection.strOpSymbol(5).Trim > 1400 And objKtbnStrc.strcSelection.strOpSymbol(5).Trim <= 1500 Then
                            '1401M~1500M
                            strOpRefKataban(UBound(strOpRefKataban)) = _
                                Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                                "1500"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                End Select

                '支持加算
                If (objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "00") Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "BBS-O" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                            objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                            objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                            objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    ElseIf objKtbnStrc.strcSelection.strSeriesKataban.Trim = "BBS-OB" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                            Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                            objKtbnStrc.strcSelection.strOpSymbol(2).Trim & CdCst.Sign.Hypen & _
                            objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1
                End If

                '付属品加算
                If objKtbnStrc.strcSelection.strOpSymbol(6).Trim.Length <> 0 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "BBS-O" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                        objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    ElseIf objKtbnStrc.strcSelection.strSeriesKataban.Trim = "BBS-OB" Then
                        strOpRefKataban(UBound(strOpRefKataban)) = _
                        Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 5) & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim & CdCst.Sign.Hypen & _
                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                    End If
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub
End Module
