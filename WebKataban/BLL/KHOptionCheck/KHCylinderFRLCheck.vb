Module KHCylinderFRLCheck

    '********************************************************************************************
    '*【関数名】
    '*  fncCheckSelectOption
    '*【処理】
    '*  シリンダチェック
    '*【概要】
    '*  シリンダＦＲＬシリーズをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Public Function fncCheckSelectOption(ByVal objKtbnStrc As KHKtbnStrc, _
                                         ByRef intKtbnStrcSeqNo As Integer, _
                                         ByRef strOptionSymbol As String, _
                                         ByRef strMessageCd As String) As Boolean

        Dim strOpArray() As String = Nothing
        Dim intLoopCnt As Integer = Nothing

        Try


            fncCheckSelectOption = True

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                'RM1003086 2010/03/30 Y.Miura 機種追加
                'Case "W3000", "W3100", "W4000", "W4100", _
                '     "R2000", "R2100", "R3000", "R3100", "R4000", "R4100", "R6000", "R6100"
                Case "W1000", "W1100", "W2000", "W2100", "W3000", "W3100", "W4000", "W4100", _
                     "R1000", "R1100", "R2000", "R2100", "R3000", "R3100", "R4000", "R4100", _
                     "R6000", "R6100", "R8000", "R8100", "W8000", "W8100"
                    '↓2013/05/20 追加
                    Dim intAsblTypePos As Integer
                    Dim intAttachPos As Integer
                    'RM1004099
                    'FRL-P4シリーズオプション追加 
                    '2010/07/27 MOD RM1007012(8月VerUP:クリーン仕様シリーズ) START --->
                    'If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" OrElse _
                    '    objKtbnStrc.strcSelection.strSeriesKataban = "R2000" OrElse _
                    '    objKtbnStrc.strcSelection.strSeriesKataban = "R2100" Then
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" OrElse _
                        objKtbnStrc.strcSelection.strKeyKataban.Trim = "P" OrElse _
                        objKtbnStrc.strcSelection.strSeriesKataban = "R2000" OrElse _
                        objKtbnStrc.strcSelection.strSeriesKataban = "R2100" Then
                        '2010/07/27 MOD RM1007012(8月VerUP:クリーン仕様シリーズ) <--- END

                        Dim bolOptionN As Boolean = False
                        Dim bolOptionT As Boolean = False
                        Dim bolOptionT6 As Boolean = False    '2017/03/09 追加 RM1702049
                        Dim bolOptionT8 As Boolean = False
                        Dim intOptionPos As Integer
                        Dim intCleanPos As Integer

                        Select Case objKtbnStrc.strcSelection.strSeriesKataban
                            Case "R2000", "R2100"
                                intOptionPos = 3
                                intCleanPos = 5
                            Case Else
                                intOptionPos = 2
                                intCleanPos = 4
                        End Select
                        'Ｐ７０クリーンルム仕様選択判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(intCleanPos).Trim
                            Case "P70", "P74"
                                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos), CdCst.Sign.Delimiter.Comma)
                                For intLoopCnt = 0 To strOpArray.Length - 1
                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case "N"
                                            bolOptionN = True
                                        Case "T"
                                            bolOptionT = True
                                            '2017/03/09 追加 RM1702049
                                        Case "T6"
                                            bolOptionT6 = True
                                        Case "T8"
                                            bolOptionT8 = True
                                    End Select
                                Next

                                'オプションでＮを選択していなかったらエラー
                                If bolOptionN = False Then
                                    intKtbnStrcSeqNo = intOptionPos
                                    strMessageCd = "W8090"
                                    fncCheckSelectOption = False
                                End If

                                'クリーン仕様Ｐ７４を選択し、オプションでＴまたはＴ８を選択していなかったらエラー
                                'T6を選択していた場合もエラーが出ないようにする  2017/03/09 追加  RM1702049
                                If objKtbnStrc.strcSelection.strOpSymbol(intCleanPos).Trim = "P74" Then
                                    If bolOptionT = False And _
                                       bolOptionT6 = False And _
                                       bolOptionT8 = False Then
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8490"
                                        fncCheckSelectOption = False
                                    End If
                                End If
                        End Select
                    End If

                    'キー形番により判定を行う
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "G" Then
                        'オプションでTまたはT8を選択したかどうか判定する
                        If InStr(objKtbnStrc.strcSelection.strOpSymbol(3), "T") = 0 Then
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W8050"
                            fncCheckSelectOption = False
                        End If

                        'RM1004099
                        'FRL-P4シリーズオプション追加
                        'ElseIf objKtbnStrc.strcSelection.strKeyKataban.Trim = "W" And _
                        '       (objKtbnStrc.strcSelection.strSeriesKataban <> "R2000" And _
                        '       objKtbnStrc.strcSelection.strSeriesKataban <> "R2100") Then
                        '二次電池対応（オプションP4,P40）の場合、T,T6,T8の指定が必須
                    ElseIf objKtbnStrc.strcSelection.strKeyKataban.Trim = "W" Then
                        Dim strOptionP4 As String = String.Empty
                        Dim strOptionT As String = String.Empty
                        Dim strOptionN As String = String.Empty     'RM1003086
                        Dim intOptionPos As Integer = 3
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "P4", "P40"
                                    strOptionP4 = strOpArray(intLoopCnt).Trim
                                Case "T", "T6", "T8"
                                    strOptionT = strOpArray(intLoopCnt).Trim
                                Case "N"    'RM1003086
                                    strOptionN = strOpArray(intLoopCnt).Trim
                            End Select
                        Next

                        'RM1004099
                        'FRL-P4シリーズオプション追加
                        'If strOptionP4 <> "" And strOptionT = "" Then
                        '    intKtbnStrcSeqNo = intOptionPos
                        '    strMessageCd = "W8790"
                        '    fncCheckSelectOption = False
                        'End If
                        ''RM1003086 
                        ''R1_00,W1_00はP4選択時にNオプションが必須
                        'Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        '    Case "W1000", "W1100", "R1000", "R1100"
                        '        If strOptionP4 <> "" And strOptionN = "" Then
                        '            intKtbnStrcSeqNo = intOptionPos
                        '            strMessageCd = "W8820"
                        '            fncCheckSelectOption = False
                        '        End If
                        'End Select
                        If strOptionP4 <> "" Then
                            'RM1003086 
                            'R1_00,W1_00はP4選択時にNオプションが必須
                            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                                Case "W1000", "W1100", "R1000", "R1100", "W8000", "W8100", "R8000", "R8100"
                                    If strOptionT = "" AndAlso strOptionN = "" Then
                                        'N,T,T6,T8の指定が必要
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8790"
                                        fncCheckSelectOption = False

                                    ElseIf strOptionT = "" Then
                                        'T,T6,T8の指定が必要
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8870"
                                        fncCheckSelectOption = False

                                    ElseIf strOptionN = "" Then
                                        'Nの指定が必要
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8820"
                                        fncCheckSelectOption = False

                                    End If

                                Case Else
                                    '上記以外、T、T6、T8を選択していない場合、エラー
                                    If strOptionT = "" Then
                                        intKtbnStrcSeqNo = intOptionPos
                                        strMessageCd = "W8870"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                        End If

                        'RM1004099
                        'FRL-P4シリーズオプション追加
                        'Else
                        '    Dim bolOptionN As Boolean = False
                        '    Dim bolOptionT As Boolean = False
                        '    Dim bolOptionT8 As Boolean = False
                        '    Dim intOptionPos As Integer
                        '    Dim intCleanPos As Integer

                        '    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        '        Case "R2000", "R2100"
                        '            intOptionPos = 3
                        '            intCleanPos = 5
                        '        Case Else
                        '            intOptionPos = 2
                        '            intCleanPos = 4
                        '    End Select
                        '    'Ｐ７０クリーンルム仕様選択判定
                        '    Select Case objKtbnStrc.strcSelection.strOpSymbol(intCleanPos).Trim
                        '        Case "P70", "P74"
                        '            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos), CdCst.Sign.Delimiter.Comma)
                        '            For intLoopCnt = 0 To strOpArray.Length - 1
                        '                Select Case strOpArray(intLoopCnt).Trim
                        '                    Case "N"
                        '                        bolOptionN = True
                        '                    Case "T"
                        '                        bolOptionT = True
                        '                    Case "T8"
                        '                        bolOptionT8 = True
                        '                End Select
                        '            Next

                        '            'オプションでＮを選択していなかったらエラー
                        '            If bolOptionN = False Then
                        '                intKtbnStrcSeqNo = intOptionPos
                        '                strMessageCd = "W8090"
                        '                fncCheckSelectOption = False
                        '            End If

                        '            'クリーン仕様Ｐ７４を選択し、オプションでＴまたはＴ８を選択していなかったらエラー
                        '            If objKtbnStrc.strcSelection.strOpSymbol(intCleanPos).Trim = "P74" Then
                        '                If bolOptionT = False And _
                        '                   bolOptionT8 = False Then
                        '                    intKtbnStrcSeqNo = intOptionPos
                        '                    strMessageCd = "W8490"
                        '                    fncCheckSelectOption = False
                        '                End If
                        '            End If
                        '    End Select
                    End If

                    '↓2013/05/20 追加
                    If objKtbnStrc.strcSelection.strKeyKataban = "W" Then
                        intAsblTypePos = 4
                        intAttachPos = 5
                    Else
                        intAsblTypePos = 3
                        intAttachPos = 4
                    End If

                    Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                        Case "W1000", "W1100", "W2000", "W2100", _
                             "W3000", "W3100", "W4000", "W4100"
                            If objKtbnStrc.strcSelection.strOpSymbol(intAsblTypePos).Trim = "U" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(intAttachPos).Trim.Length = 0 Then
                                    intKtbnStrcSeqNo = intAttachPos
                                    strMessageCd = "W8060"
                                    fncCheckSelectOption = False
                                End If
                            End If
                    End Select
                    '↑2013/05/20 追加

                Case "W8000", "W8100", "R8000", "R8100"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "G" Then
                        'オプションでTまたはT8を選択したかどうか判定する
                        If InStr(objKtbnStrc.strcSelection.strOpSymbol(3), "T") = 0 Then
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W8050"
                            fncCheckSelectOption = False
                        End If
                    End If

                Case "B7019"
                    Dim bolOptionM As Boolean = False
                    Dim bolOptionO As Boolean = False

                    'OP分解
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(2), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "M", "MG"
                                bolOptionM = True
                            Case "G"
                                If bolOptionM = True Then
                                    If bolOptionO = False Then
                                        intKtbnStrcSeqNo = 2
                                        strMessageCd = "W8070"
                                        fncCheckSelectOption = False
                                    End If
                                End If
                            Case "-G"
                                If bolOptionM = True Then
                                    If bolOptionO = True Then
                                        intKtbnStrcSeqNo = 2
                                        strMessageCd = "W8080"
                                        fncCheckSelectOption = False
                                    End If
                                Else
                                    intKtbnStrcSeqNo = 2
                                    strMessageCd = "W8080"
                                    fncCheckSelectOption = False
                                End If
                            Case Else
                                If bolOptionM = True Then
                                    bolOptionO = True
                                End If
                        End Select
                    Next
                    '2011/06/20 RM1106028(7月VerUP:7080,A7070シリーズ) START--->
                Case "7080", "A7070"
                    Dim bolOptionM As Boolean = False

                    'OP分解
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(2), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "M", "MG"
                                bolOptionM = True
                        End Select
                    Next

                    'OP分解
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "-G"
                                If bolOptionM = False Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W8080"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    Next
                    '2011/06/20 RM1106028(7月VerUP:7080,A7070シリーズ) <---END

                    'RM0904032 2009/06/18 Y.Miura
                    'Case "C1000", "C1010", "C1020", "C1030", "C1040", _
                    '     "C1050", "C1060", "C2500", "C2520", "C2530", _
                    '     "C2550", "C3000", "C3010", "C3020", "C3030", _
                    '     "C3040", "C3050", "C3060", "C3070", "C4000", _
                    '     "C4010", "C4020", "C4030", "C4040", "C4050", _
                    '     "C4060", "C4070", "C6020", "C6030", "C6050", _
                    '     "C6060", "C6070", "C6500", "C8000", "C8010", _
                    '     "C8020", "C8030", "C8040", "C8050", "C8060", _
                    '     "C8070"

                    '↓2013/05/20 W%000シリーズ追加
                    '↓RM1308014 2013/08/05 追加
                Case "C1000", "C1010", "C1020", "C1030", "C1040", "C1050", "C1060", _
                     "C2000", "C2010", "C2020", "C2030", "C2040", "C2050", "C2060", _
                     "C2500", "C2520", "C2530", "C2550", _
                     "C3000", "C3010", "C3020", "C3030", "C3040", "C3050", "C3060", "C3070", _
                     "C4000", "C4010", "C4020", "C4030", "C4040", "C4050", "C4060", "C4070", _
                     "C6020", "C6030", "C6050", "C6060", "C6070", "C6500", _
                     "C8000", "C8010", "C8020", "C8030", "C8040", "C8050", "C8060", "C8070", _
                     "W1000    W", "W2000    W", "W3000    W", "W4000    W", _
                     "W1100    W", "W2100    W", "W3100    W", "W4100    W", _
                     "L3000", "L4000", "L8000"

                    Dim intAsblTypePos As Integer
                    Dim intAttachPos As Integer
                    If objKtbnStrc.strcSelection.strKeyKataban = "W" Then
                        intAsblTypePos = 4
                        intAttachPos = 5

                        '2010/10/29 RM1011020(12月VerUP:FRLシリーズ) START--->
                        '↓RM1308014 2013/08/05 追加
                        Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                            Case "C2000", "C2010", "C2500", "C3000", "C3010", _
                                 "C4000", "C4010", "C6500", "C8000", "C8010", _
                                 "L3000", "L4000", "L8000"

                                'OP分解
                                Dim bolOptionM As Boolean = False
                                Dim bolOptionFlg As Boolean = False
                                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(3), CdCst.Sign.Delimiter.Comma)
                                For intLoopCnt = 0 To strOpArray.Length - 1
                                    Select Case strOpArray(intLoopCnt).Trim
                                        Case "M1"
                                            bolOptionM = True
                                            'Case "C", "F", "F1"
                                        Case "C"
                                            bolOptionFlg = True
                                    End Select
                                Next

                                'オプション「M1」を指定した場合
                                If bolOptionM Then
                                    'オプション「C,F,F1」のいずれかを指定
                                    If Not bolOptionFlg Then
                                        'エラーメッセージ「オプション「M1」選択時は、ドレン排出オプション「C」「F」「F1」のいずれかを選定してください。」
                                        intKtbnStrcSeqNo = 3
                                        strMessageCd = "W0840"
                                        fncCheckSelectOption = False

                                    End If
                                End If
                        End Select

                        '2010/10/29 RM1011020(12月VerUP:FRLシリーズ) <---END
                    Else
                        intAsblTypePos = 3
                        intAttachPos = 4
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(intAsblTypePos).Trim = "U" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(intAttachPos).Trim.Length = 0 Then
                            intKtbnStrcSeqNo = intAttachPos
                            strMessageCd = "W8060"
                            fncCheckSelectOption = False
                        End If
                    End If

                Case "RN3000", "RN4000", "RN8000"
                    'オプションでTまたはT8を選択したかどうか判定する
                    If objKtbnStrc.strcSelection.strOpSymbol(2).IndexOf("T") >= 0 Then
                    Else
                        intKtbnStrcSeqNo = 2
                        strMessageCd = "W8050"
                        fncCheckSelectOption = False
                    End If
                Case "CXU10"
                    'オプションを1つも選択していない、または「X」のみを選択している場合はエラー
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        'Case "1"   'RM1003086 追加
                        Case "1", "5"
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim.Length = 0 Or _
                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "X" Then
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W8610"
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "CXU30"
                    'オプションを1つも選択していない、または「X」のみを選択している場合はエラー
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "1"
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim.Length = 0 Or _
                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "X" Then
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W8610"
                                fncCheckSelectOption = False
                            End If
                        Case "2"
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim.Length = 0 Or _
                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "X" Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W8610"
                                fncCheckSelectOption = False
                            End If
                    End Select

                Case "RB500"
                    'RM1004012 2010/04/23 Y.Miura
                    'P4選択時、オプションで「N」かつ「T」を選択していない場合は、エラー表示する
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        '2010/08/24 MOD RM1008009(9月VerUP:RB500シリーズ 機種追加) START --->
                        'Case "1"
                        '    Dim bolOptionN As Boolean = False
                        '    Dim bolOptionT As Boolean = False
                        Case "4"
                            Dim bolOption As Boolean = False
                            '2010/08/24 MOD RM1008009(9月VerUP:RB500シリーズ 機種追加) <--- END

                            'OP分解
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(4), CdCst.Sign.Delimiter.Comma)
                            '2010/08/24 MOD RM1008009(9月VerUP:RB500シリーズ 機種追加) START --->
                            'For intLoopCnt = 0 To strOpArray.Length - 1
                            '    Select Case strOpArray(intLoopCnt).Trim
                            '        Case "N"
                            '            bolOptionN = True
                            '        Case "T"
                            '            bolOptionT = True
                            '    End Select
                            'Next
                            If strOpArray.Length >= 2 AndAlso _
                                strOpArray(0).Equals("N") AndAlso strOpArray(1).Equals("T") Then
                                bolOption = True
                            End If
                            '2010/08/24 MOD RM1008009(9月VerUP:RB500シリーズ 機種追加) <--- END
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(6).Trim
                                Case "P4"
                                    '2010/08/24 MOD RM1008009(9月VerUP:RB500シリーズ 機種追加) START --->
                                    'If bolOptionN And bolOptionT Then
                                    If bolOption Then
                                        '2010/08/24 MOD RM1008009(9月VerUP:RB500シリーズ 機種追加) <--- END
                                    Else
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W8840"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                    End Select

            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function
End Module
