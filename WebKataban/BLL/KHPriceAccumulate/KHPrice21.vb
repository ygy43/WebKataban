'************************************************************************************
'*  ProgramID  ：KHPrice21
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/05   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：マイクロエレッサ
'*
'************************************************************************************
Module KHPrice21

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim strOpArray() As String
        Dim intLoopCnt1 As Integer
        Dim intLoopCnt2 As Integer

        Dim bolOptionFlg1 As Boolean
        Dim bolOptionFlg2 As Boolean
        Dim bolOptionFlg3 As Boolean
        Dim bolOptionFlg4 As Boolean
        Dim bolOptionFlg5 As Boolean
        Dim bolOptionFlg6 As Boolean
        Dim bolOptionFlg7 As Boolean

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '基本価格キー
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "1137", "2001", "2415", "3502", "2215", "2401", "2216"
                    Select Case Mid(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2, 1)
                        Case "", "C"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2)
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select
                Case "2302"
                    If objKtbnStrc.strcSelection.strKeyKataban = "W" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "-W"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                Case "2303", "2304"
                    If objKtbnStrc.strcSelection.strKeyKataban = "W" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1).Trim & "-W"
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        Select Case Mid(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2, 1)
                            Case "", "C"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 1)
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           Left(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, 2)
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                Case Else
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                        Case "1126", "1226", "1326", "1226J"
                            If objKtbnStrc.strcSelection.strKeyKataban = "W" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1).Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            Else
                                'RM1304XXX 2014/03/14
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        Case Else
                            'RM1304XXX 2014/03/14
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim
                            decOpAmount(UBound(decOpAmount)) = 1
                    End Select

            End Select

                    'クリーン仕様時に基本価格キーを変更
                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                        Case "1019", "1144", "1219", "1244", "2619"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "P80", "P90", "P94"
                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & "-" & objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            End Select
                    End Select

                    'オプション加算価格キー
            For intLoopCnt1 = 2 To objKtbnStrc.strcSelection.strOpSymbol.Length - 1
                'RM1304XXX 2014/03/14
                '2014/04/24修正
                If objKtbnStrc.strcSelection.strSeriesKataban.Trim <> "2302" Then
                    'If objKtbnStrc.strcSelection.strKeyKataban <> "W" Then
                    If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt1).Trim <> "" Then
                        bolOptionFlg1 = False
                        bolOptionFlg2 = False
                        bolOptionFlg3 = False
                        bolOptionFlg4 = False
                        bolOptionFlg5 = False
                        bolOptionFlg6 = False
                        bolOptionFlg7 = False

                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt1), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt2 = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt2).Trim
                                Case ""
                                Case "P80", "P90", "P94"
                                Case Else
                                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                        Case "1137", "1237", "7070", "7170", "1144", "3500", "7080"
                                            Select Case strOpArray(intLoopCnt2).Trim
                                                Case "Y"
                                                    bolOptionFlg1 = True
                                                Case "F"
                                                    bolOptionFlg2 = True
                                                Case "F1"
                                                    bolOptionFlg2 = True
                                                    bolOptionFlg3 = True
                                                Case "X"
                                                    bolOptionFlg4 = True
                                                Case "J"
                                                    bolOptionFlg5 = True
                                                Case "EJ"
                                                    bolOptionFlg5 = True
                                                    bolOptionFlg6 = True
                                                Case "FJ"
                                                    bolOptionFlg5 = True
                                                    bolOptionFlg7 = True
                                                Case "F1J"
                                                    bolOptionFlg3 = True
                                                    bolOptionFlg5 = True
                                                    bolOptionFlg7 = True
                                            End Select
                                    End Select

                                    'オプションキーセット
                                    Select Case True
                                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1237" And strOpArray(intLoopCnt2).Trim = "J" And bolOptionFlg5 = True
                                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1237" And strOpArray(intLoopCnt2).Trim = "EJ" And bolOptionFlg6 = True
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "E"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1237" And strOpArray(intLoopCnt2).Trim = "FJ" And bolOptionFlg7 = True
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "F"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1237" And strOpArray(intLoopCnt2).Trim = "F1J" And bolOptionFlg7 = True
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "F1"
                                            decOpAmount(UBound(decOpAmount)) = 1
                                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1137" And strOpArray(intLoopCnt2).Trim = "F" And bolOptionFlg2 = True
                                        Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1137" And strOpArray(intLoopCnt2).Trim = "F1" And bolOptionFlg3 = True
                                        Case Else
                                            If bolOptionFlg5 = True Then
                                                Select Case strOpArray(intLoopCnt2).Trim
                                                    Case "Z", "M", "MG", "MG2"
                                                        If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1237" Then
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                                       strOpArray(intLoopCnt2).Trim & "(J)"
                                                            decOpAmount(UBound(decOpAmount)) = 1
                                                        Else
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                                       strOpArray(intLoopCnt2).Trim & "J"
                                                            decOpAmount(UBound(decOpAmount)) = 1
                                                        End If
                                                    Case Else
                                                        If strOpArray(intLoopCnt2).Trim = "-G" Then
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "G"
                                                            decOpAmount(UBound(decOpAmount)) = 1
                                                        Else
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                                       strOpArray(intLoopCnt2).Trim
                                                            decOpAmount(UBound(decOpAmount)) = 1
                                                        End If
                                                End Select
                                            Else
                                                If strOpArray(intLoopCnt2).Trim = "-G" Then
                                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "G"
                                                    decOpAmount(UBound(decOpAmount)) = 1
                                                Else
                                                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                                        Case "1126", "1226", "1326", "1226J"
                                                            If objKtbnStrc.strcSelection.strKeyKataban = "W" Then
                                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & "-W-" & _
                                                                                                           strOpArray(intLoopCnt2).Trim
                                                                decOpAmount(UBound(decOpAmount)) = 1
                                                            Else
                                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                                           strOpArray(intLoopCnt2).Trim
                                                                decOpAmount(UBound(decOpAmount)) = 1
                                                            End If
                                                        Case Else
                                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                                       strOpArray(intLoopCnt2).Trim
                                                            decOpAmount(UBound(decOpAmount)) = 1
                                                    End Select
                                                    
                                                End If
                                            End If
                                    End Select

                                    'クリーン仕様時にオプション加算価格キーを変更
                                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                        Case "1019", "1144", "1219", "1244", "2619"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "P80", "P90", "P94"
                                                    strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            End Select
                                    End Select
                            End Select
                        Next

                        Select Case True
                            Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1137" And bolOptionFlg1 = True And bolOptionFlg2 = True And bolOptionFlg3 = True
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "F1Y"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1137" And bolOptionFlg1 = True And bolOptionFlg2 = True And bolOptionFlg3 = False
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "FY"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1137" And bolOptionFlg1 = False And bolOptionFlg2 = True And bolOptionFlg3 = True
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "F1"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1137" And bolOptionFlg1 = False And bolOptionFlg2 = True And bolOptionFlg3 = False
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "F"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1237" And bolOptionFlg4 = True And bolOptionFlg5 = True
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "J(X)"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1237" And bolOptionFlg4 = False And bolOptionFlg5 = True
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "J"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                    End If
                    'End If
                Else
                    '2302シリーズのみ
                    If objKtbnStrc.strcSelection.strKeyKataban = "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt1).Trim <> "" Then
                            bolOptionFlg1 = False
                            bolOptionFlg2 = False
                            bolOptionFlg3 = False
                            bolOptionFlg4 = False
                            bolOptionFlg5 = False
                            bolOptionFlg6 = False
                            bolOptionFlg7 = False

                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt1), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt2 = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt2).Trim
                                    Case ""
                                    Case "P80", "P90", "P94"
                                    Case Else
                                        Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                            Case "1137", "1237", "7070", "7170", "1144", "3500", "7080"
                                                Select Case strOpArray(intLoopCnt2).Trim
                                                    Case "Y"
                                                        bolOptionFlg1 = True
                                                    Case "F"
                                                        bolOptionFlg2 = True
                                                    Case "F1"
                                                        bolOptionFlg2 = True
                                                        bolOptionFlg3 = True
                                                    Case "X"
                                                        bolOptionFlg4 = True
                                                    Case "J"
                                                        bolOptionFlg5 = True
                                                    Case "EJ"
                                                        bolOptionFlg5 = True
                                                        bolOptionFlg6 = True
                                                    Case "FJ"
                                                        bolOptionFlg5 = True
                                                        bolOptionFlg7 = True
                                                    Case "F1J"
                                                        bolOptionFlg3 = True
                                                        bolOptionFlg5 = True
                                                        bolOptionFlg7 = True
                                                End Select
                                        End Select

                                        'オプションキーセット
                                        Select Case True
                                            Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1237" And strOpArray(intLoopCnt2).Trim = "J" And bolOptionFlg5 = True
                                            Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1237" And strOpArray(intLoopCnt2).Trim = "EJ" And bolOptionFlg6 = True
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "E"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1237" And strOpArray(intLoopCnt2).Trim = "FJ" And bolOptionFlg7 = True
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "F"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1237" And strOpArray(intLoopCnt2).Trim = "F1J" And bolOptionFlg7 = True
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "F1"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                            Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1137" And strOpArray(intLoopCnt2).Trim = "F" And bolOptionFlg2 = True
                                            Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1137" And strOpArray(intLoopCnt2).Trim = "F1" And bolOptionFlg3 = True
                                            Case Else
                                                If bolOptionFlg5 = True Then
                                                    Select Case strOpArray(intLoopCnt2).Trim
                                                        Case "Z", "M", "MG", "MG2"
                                                            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1237" Then
                                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                                           strOpArray(intLoopCnt2).Trim & "(J)"
                                                                decOpAmount(UBound(decOpAmount)) = 1
                                                            Else
                                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                                           strOpArray(intLoopCnt2).Trim & "J"
                                                                decOpAmount(UBound(decOpAmount)) = 1
                                                            End If
                                                        Case Else
                                                            If strOpArray(intLoopCnt2).Trim = "-G" Then
                                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "G"
                                                                decOpAmount(UBound(decOpAmount)) = 1
                                                            Else
                                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                                           strOpArray(intLoopCnt2).Trim
                                                                decOpAmount(UBound(decOpAmount)) = 1
                                                            End If
                                                    End Select
                                                Else
                                                    If strOpArray(intLoopCnt2).Trim = "-G" Then
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "G"
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    Else
                                                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                                                   strOpArray(intLoopCnt2).Trim
                                                        decOpAmount(UBound(decOpAmount)) = 1
                                                    End If
                                                End If
                                        End Select

                                        'クリーン仕様時にオプション加算価格キーを変更
                                        Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                                            Case "1019", "1144", "1219", "1244", "2619"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                    Case "P80", "P90", "P94"
                                                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                End Select
                                        End Select
                                End Select
                            Next

                            Select Case True
                                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1137" And bolOptionFlg1 = True And bolOptionFlg2 = True And bolOptionFlg3 = True
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "F1Y"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1137" And bolOptionFlg1 = True And bolOptionFlg2 = True And bolOptionFlg3 = False
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "FY"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1137" And bolOptionFlg1 = False And bolOptionFlg2 = True And bolOptionFlg3 = True
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "F1"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1137" And bolOptionFlg1 = False And bolOptionFlg2 = True And bolOptionFlg3 = False
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "F"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1237" And bolOptionFlg4 = True And bolOptionFlg5 = True
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "J(X)"
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Case objKtbnStrc.strcSelection.strSeriesKataban.Trim = "1237" And bolOptionFlg4 = False And bolOptionFlg5 = True
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & "J"
                                    decOpAmount(UBound(decOpAmount)) = 1
                            End Select
                        End If
                    End If
                End If
            Next

            'RM1304XXX 2014/03/14
            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "2302" Then
                If objKtbnStrc.strcSelection.strKeyKataban = "W" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) & "-W-" & objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    decOpAmount(UBound(decOpAmount)) = 1
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
