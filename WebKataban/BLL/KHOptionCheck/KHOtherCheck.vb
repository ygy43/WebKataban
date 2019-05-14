Module KHOtherCheck

    '********************************************************************************************
    '*【関数名】
    '*  fncCheckSelectOption
    '*【処理】
    '*  その他チェック
    '*【概要】
    '*  その他の機種をチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2007/12/20      更新者：NII A.Takahashi
    '*  ・FRL難燃シリーズ-G4(W3*00-G4/W4*00-G4/W8*00-G4/R3*00-G4/R4*00-G4/R8*00-G4)シリーズ
    '*  　オプション「T」「T8」を選択しないとエラー表示するよう修正
    '*
    '*  ・受付No：RM0904032  FRL2000新発売
    '*                                          更新日：2009/06/18      更新者：Y.Miura
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
                Case "HD"
                    If Not objKtbnStrc.strcSelection.strOpSymbol(2).Contains("G") Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Case "AC100V", "AC200V"
                            Case Else
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W8020"
                                fncCheckSelectOption = False
                        End Select
                    Else
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Case "AC100V", "AC200V", "AC110V", "AC115V", "AC120V", "AC127V", "AC208V", "AC220V", _
                                "AC230V", "AC240V", "AC380V", "AC400V", "AC415V", "AC440V", "AC460V", _
                                "AC480V"
                            Case Else
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W8020"
                                fncCheckSelectOption = False
                        End Select
                    End If
                Case "ESM"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1) = "B" Then
                            'ベルト長さは5桁表示
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Length <> 5 Then
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W9090"
                                fncCheckSelectOption = False
                            End If

                            'ベルト長さは２mm単位
                            Dim intMod As Integer = 0
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Length <> 0 Then
                                intMod = objKtbnStrc.strcSelection.strOpSymbol(3) Mod 2
                                If intMod <> 0 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W9100"
                                    fncCheckSelectOption = False
                                End If

                            End If
                        End If
                    End If
                Case "ECV"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "F" Or _
                         objKtbnStrc.strcSelection.strKeyKataban.Trim = "X" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "Y" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3) <> Nothing Then
                            'ストロークは3桁表示
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Length <> 3 Then
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W9110"
                                fncCheckSelectOption = False
                            End If

                            'ストロークは50mm単位
                            Dim intMod As Integer = 0
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Length <> 0 Then
                                intMod = objKtbnStrc.strcSelection.strOpSymbol(3) Mod 5
                                If intMod <> 0 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0330"
                                    fncCheckSelectOption = False
                                End If

                            End If
                        End If
                    End If

                Case "VNA"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1) = "32" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(1) = "40" Then
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Contains("H") And _
                                 objKtbnStrc.strcSelection.strOpSymbol(3).Contains("L") Then

                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W0950"
                                fncCheckSelectOption = False

                            End If
                        End If
                    End If
                Case "RGIS", "RGOS"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "B" Or objKtbnStrc.strcSelection.strKeyKataban.Trim = "Y" Then
                        'オプションNNN判定 2014/07/25
                        If objKtbnStrc.strcSelection.strOpSymbol(8) = "N" And _
                        objKtbnStrc.strcSelection.strOpSymbol(9) = "N" And _
                        objKtbnStrc.strcSelection.strOpSymbol(10) = "N" Then
                            intKtbnStrcSeqNo = 8
                            strMessageCd = "W9040"
                            fncCheckSelectOption = False
                        Else
                        End If
                    End If
                Case "RGIT", "RGCT", "RGOL", "PCIS", "PCOS"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "B" Then
                        'オプションNNN判定 2014/07/25
                        If objKtbnStrc.strcSelection.strOpSymbol(8) = "N" And _
                        objKtbnStrc.strcSelection.strOpSymbol(9) = "N" And _
                        objKtbnStrc.strcSelection.strOpSymbol(10) = "N" Then
                            intKtbnStrcSeqNo = 8
                            strMessageCd = "W9040"
                            fncCheckSelectOption = False
                        Else
                        End If
                    End If
                Case "RGCS", "RGIL"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "Y" Then
                        'オプションNNN判定 2014/07/25
                        If objKtbnStrc.strcSelection.strOpSymbol(8) = "N" And _
                        objKtbnStrc.strcSelection.strOpSymbol(9) = "N" And _
                        objKtbnStrc.strcSelection.strOpSymbol(10) = "N" Then
                            intKtbnStrcSeqNo = 8
                            strMessageCd = "W9040"
                            fncCheckSelectOption = False
                        Else
                        End If
                    End If
                Case "RGIM", "RGCM", "RGID", "RGCD", "RGIB"
                    'オプションNNN判定 2015/08/18
                    If objKtbnStrc.strcSelection.strOpSymbol(8) = "N" And _
                    objKtbnStrc.strcSelection.strOpSymbol(9) = "N" And _
                    objKtbnStrc.strcSelection.strOpSymbol(10) = "N" Then
                        intKtbnStrcSeqNo = 8
                        strMessageCd = "W9040"
                        fncCheckSelectOption = False
                    Else
                    End If
                Case "VPCM", "VPR2M"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "2" Then
                        'ストローク範囲判定
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(10).Trim) = CInt(objKtbnStrc.strcSelection.strOpSymbol(11).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) Then
                        Else
                            intKtbnStrcSeqNo = 11
                            strMessageCd = "W8130"
                            fncCheckSelectOption = False
                        End If
                    End If
                Case "MXKML"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                        Case "1"
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "0" Then
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W8140"
                                fncCheckSelectOption = False
                            End If
                        Case "2"
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "0" And _
                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "0" Then
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W8140"
                                fncCheckSelectOption = False
                            End If
                        Case "3"
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "0" And _
                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "0" And _
                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "0" Then
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W8140"
                                fncCheckSelectOption = False
                            End If
                        Case "4"
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "0" And _
                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "0" And _
                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "0" And _
                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "0" Then
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W8140"
                                fncCheckSelectOption = False
                            End If
                        Case "5"
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "0" And _
                               objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "0" And _
                               objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "0" And _
                               objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "0" And _
                               objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "0" Then
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W8140"
                                fncCheckSelectOption = False
                            End If
                    End Select

                    '混合搭載時のKML50とKML60が混載判定
                    Dim bolOption1 As Boolean = False
                    Dim bolOption4 As Boolean = False
                    Dim bolOption6 As Boolean = False

                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "00-0" Then
                    Else
                        '混合搭載選択時
                        For intLoopCnt = 3 To objKtbnStrc.strcSelection.strOpSymbol.Length - 1
                            If intLoopCnt <= 7 Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt).Trim
                                    Case "1"
                                        bolOption1 = True
                                    Case "4"
                                        bolOption4 = True
                                    Case "6"
                                        bolOption6 = True
                                End Select
                            End If
                        Next

                        Select Case True
                            Case bolOption1 = True And bolOption4 = True
                            Case bolOption1 = True And bolOption6 = True
                            Case Else
                                intKtbnStrcSeqNo = 3
                                strMessageCd = "W8150"
                                fncCheckSelectOption = False
                        End Select
                    End If
                Case "RV3SA"
                    '角度指定なし時（0）はチェック不要
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "0" Then
                        '揺動角度判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Case "45"
                                '指定角度の範囲判定
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) >= 30 And _
                                   CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) <= 270 Then
                                Else
                                    intKtbnStrcSeqNo = 2
                                    strMessageCd = "W8110"
                                    fncCheckSelectOption = False
                                End If
                            Case "90"
                                '指定角度の範囲判定
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) >= 30 And _
                                   CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) <= 180 Then
                                Else
                                    intKtbnStrcSeqNo = 2
                                    strMessageCd = "W8120"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    End If
                Case "RV3DA"
                    '角度指定なし時（0）はチェック不要
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "0" Then
                        '揺動角度判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Case "45"
                                '指定角度の範囲判定
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) >= 30 And _
                                   CInt(objKtbnStrc.strcSelection.strOpSymbol(2).Trim) <= 90 Then
                                Else
                                    intKtbnStrcSeqNo = 2
                                    strMessageCd = "W8100"
                                    fncCheckSelectOption = False
                                End If
                            Case "90"
                                '設定なし
                        End Select
                    End If
                Case "RGCT", "RGIT"
                    '軸間距離判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "063", "080", "110"
                        Case Else
                    End Select

                    '2009/08/28 Y.Miura
                    'KHCylinderFRLCheckに移動
                    'Case "W3000", "W3100", "W4000", "W4100", _
                    '     "R2000", "R2100", "R3000", "R3100", "R4000", "R4100", "R6000", "R6100"
                    '    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "G" Then
                    '        'オプションでTまたはT8を選択したかどうか判定する
                    '        If InStr(objKtbnStrc.strcSelection.strOpSymbol(3), "T") = 0 Then
                    '            intKtbnStrcSeqNo = 3
                    '            strMessageCd = "W8050"
                    '            fncCheckSelectOption = False
                    '        End If
                    '    ElseIf objKtbnStrc.strcSelection.strKeyKataban.Trim = "W" And _
                    '           (objKtbnStrc.strcSelection.strSeriesKataban <> "R2000" And _
                    '           objKtbnStrc.strcSelection.strSeriesKataban <> "R2100") Then
                    '    Else
                    '        Dim bolOptionN As Boolean = False
                    '        Dim bolOptionT As Boolean = False
                    '        Dim bolOptionT8 As Boolean = False
                    '        Dim intOptionPos As Integer
                    '        Dim intCleanPos As Integer

                    '        Select Case objKtbnStrc.strcSelection.strSeriesKataban
                    '            Case "R2000", "R2100"
                    '                intOptionPos = 3
                    '                intCleanPos = 5
                    '            Case Else
                    '                intOptionPos = 2
                    '                intCleanPos = 4
                    '        End Select
                    '        'Ｐ７０クリーンルム仕様選択判定
                    '        Select Case objKtbnStrc.strcSelection.strOpSymbol(intCleanPos).Trim
                    '            Case "P70", "P74"
                    '                strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos), CdCst.Sign.Delimiter.Comma)
                    '                For intLoopCnt = 0 To strOpArray.Length - 1
                    '                    Select Case strOpArray(intLoopCnt).Trim
                    '                        Case "N"
                    '                            bolOptionN = True
                    '                        Case "T"
                    '                            bolOptionT = True
                    '                        Case "T8"
                    '                            bolOptionT8 = True
                    '                    End Select
                    '                Next

                    '                'オプションでＮを選択していなかったらエラー
                    '                If bolOptionN = False Then
                    '                    intKtbnStrcSeqNo = intOptionPos
                    '                    strMessageCd = "W8090"
                    '                    fncCheckSelectOption = False
                    '                End If

                    '                'クリーン仕様Ｐ７４を選択し、オプションでＴまたはＴ８を選択していなかったらエラー
                    '                If objKtbnStrc.strcSelection.strOpSymbol(intCleanPos).Trim = "P74" Then
                    '                    If bolOptionT = False And _
                    '                       bolOptionT8 = False Then
                    '                        intKtbnStrcSeqNo = intOptionPos
                    '                        strMessageCd = "W8490"
                    '                        fncCheckSelectOption = False
                    '                    End If
                    '                End If
                    '        End Select
                    '    End If
                    'Case "W8000", "W8100", "R8000", "R8100"
                    '    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "G" Then
                    '        'オプションでTまたはT8を選択したかどうか判定する
                    '        If InStr(objKtbnStrc.strcSelection.strOpSymbol(3), "T") = 0 Then
                    '            intKtbnStrcSeqNo = 3
                    '            strMessageCd = "W8050"
                    '            fncCheckSelectOption = False
                    '        End If
                    '    End If
                    'Case "7080", "A7070", "B7019"
                    '    Dim bolOptionM As Boolean = False
                    '    Dim bolOptionO As Boolean = False

                    '    'OP分解
                    '    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(2), CdCst.Sign.Delimiter.Comma)
                    '    For intLoopCnt = 0 To strOpArray.Length - 1
                    '        Select Case strOpArray(intLoopCnt).Trim
                    '            Case ""
                    '            Case "M", "MG"
                    '                bolOptionM = True
                    '            Case "G"
                    '                If bolOptionM = True Then
                    '                    If bolOptionO = False Then
                    '                        intKtbnStrcSeqNo = 2
                    '                        strMessageCd = "W8070"
                    '                        fncCheckSelectOption = False
                    '                    End If
                    '                End If
                    '            Case "-G"
                    '                If bolOptionM = True Then
                    '                    If bolOptionO = True Then
                    '                        intKtbnStrcSeqNo = 2
                    '                        strMessageCd = "W8080"
                    '                        fncCheckSelectOption = False
                    '                    End If
                    '                Else
                    '                    intKtbnStrcSeqNo = 2
                    '                    strMessageCd = "W8080"
                    '                    fncCheckSelectOption = False
                    '                End If
                    '            Case Else
                    '                If bolOptionM = True Then
                    '                    bolOptionO = True
                    '                End If
                    '        End Select
                    '    Next
                    '    'RM0904032 2009/06/18 Y.Miura
                    '    'Case "C1000", "C1010", "C1020", "C1030", "C1040", _
                    '    '     "C1050", "C1060", "C2500", "C2520", "C2530", _
                    '    '     "C2550", "C3000", "C3010", "C3020", "C3030", _
                    '    '     "C3040", "C3050", "C3060", "C3070", "C4000", _
                    '    '     "C4010", "C4020", "C4030", "C4040", "C4050", _
                    '    '     "C4060", "C4070", "C6020", "C6030", "C6050", _
                    '    '     "C6060", "C6070", "C6500", "C8000", "C8010", _
                    '    '     "C8020", "C8030", "C8040", "C8050", "C8060", _
                    '    '     "C8070"
                    'Case "C1000", "C1010", "C1020", "C1030", "C1040", "C1050", "C1060", _
                    '     "C2000", "C2010", "C2020", "C2030", "C2040", "C2050", "C2060", _
                    '     "C2500", "C2520", "C2530", "C2550", _
                    '     "C3000", "C3010", "C3020", "C3030", "C3040", "C3050", "C3060", "C3070", _
                    '     "C4000", "C4010", "C4020", "C4030", "C4040", "C4050", "C4060", "C4070", _
                    '     "C6020", "C6030", "C6050", "C6060", "C6070", "C6500", _
                    '     "C8000", "C8010", "C8020", "C8030", "C8040", "C8050", "C8060", "C8070"
                    '    Dim intAsblTypePos As Integer
                    '    Dim intAttachPos As Integer
                    '    If objKtbnStrc.strcSelection.strKeyKataban = "W" Then
                    '        intAsblTypePos = 4
                    '        intAttachPos = 5
                    '    Else
                    '        intAsblTypePos = 3
                    '        intAttachPos = 4
                    '    End If
                    '    If objKtbnStrc.strcSelection.strOpSymbol(intAsblTypePos).Trim = "U" Then
                    '        If objKtbnStrc.strcSelection.strOpSymbol(intAttachPos).Trim.Length = 0 Then
                    '            intKtbnStrcSeqNo = intAttachPos
                    '            strMessageCd = "W8060"
                    '            fncCheckSelectOption = False
                    '        End If
                    '    End If
                    'Case "RN3000", "RN4000", "RN8000"
                    '    'オプションでTまたはT8を選択したかどうか判定する
                    '    If objKtbnStrc.strcSelection.strOpSymbol(2).IndexOf("T") >= 0 Then
                    '    Else
                    '        intKtbnStrcSeqNo = 2
                    '        strMessageCd = "W8050"
                    '        fncCheckSelectOption = False
                    '    End If
                    'Case "CXU10"
                    '    'オプションを1つも選択していない、または「X」のみを選択している場合はエラー
                    '    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    '        Case "1"
                    '            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim.Length = 0 Or _
                    '               objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "X" Then
                    '                intKtbnStrcSeqNo = 3
                    '                strMessageCd = "W8610"
                    '                fncCheckSelectOption = False
                    '            End If
                    '    End Select
                    'Case "CXU30"
                    '    'オプションを1つも選択していない、または「X」のみを選択している場合はエラー
                    '    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    '        Case "1"
                    '            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim.Length = 0 Or _
                    '               objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "X" Then
                    '                intKtbnStrcSeqNo = 3
                    '                strMessageCd = "W8610"
                    '                fncCheckSelectOption = False
                    '            End If
                    '        Case "2"
                    '            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim.Length = 0 Or _
                    '               objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "X" Then
                    '                intKtbnStrcSeqNo = 7
                    '                strMessageCd = "W8610"
                    '                fncCheckSelectOption = False
                    '            End If
                    '    End Select
                    '2010/07/28 ADD RM1007012(8月VerUP：SHDシリーズ) START --->
                Case "SHD3"
                    'オプションでEまたはE1またはE2を選択したかどうか判定する
                    If InStr(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, "E") = 0 Then
                        intKtbnStrcSeqNo = 5
                        strMessageCd = "W2760"
                        fncCheckSelectOption = False
                    End If
                    '2010/07/28 ADD RM1007012(8月VerUP：SHDシリーズ) <--- END

                Case "FX1004", "FX1011", "FX1037"
                    '201501月次更新
                    ''オプション「M1」を指定した場合
                    'If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "M1" Then
                    '    'オプション「C,F,F1」のいずれかを指定
                    '    If Not objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "C" Then
                    '        'エラーメッセージ「オプション「M1」選択時は、ドレン排出オプション「C」を選定してください。」
                    '        intKtbnStrcSeqNo = 4
                    '        strMessageCd = "W0840"
                    '        fncCheckSelectOption = False

                    '    End If
                    'End If
                Case "CAC4"
                    If objKtbnStrc.strcSelection.strKeyKataban = "S" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(14).Contains("Y1") = False Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W9070"
                            fncCheckSelectOption = False
                        End If
                    End If
            End Select

        Catch ex As Exception

            fncCheckSelectOption = False

            Throw ex

        End Try

    End Function

End Module
