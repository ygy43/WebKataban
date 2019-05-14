Module KHCylinderCheck

    '********************************************************************************************
    '*【関数名】
    '*  fncCheckSelectOption
    '*【処理】
    '*  シリンダチェック
    '*【概要】
    '*  シリンダをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2007/05/16      更新者：NII A.Takahashi
    '*  ・T2W/T3Wスイッチ追加に伴い、ストロークチェックロジックを修正
    '*  ・受付No：RM0906034  二次電池対応機器対応
    '*                                      更新日：2009/08/05   更新者：Y.Miura
    '*  ・受付No：RM1112XXX  SMGの5mm毎ストロークチェック
    '*                                      更新日：2011/12/22   更新者：Y.Tachi
    '********************************************************************************************
    Public Function fncCheckSelectOption(ByVal objKtbnStrc As KHKtbnStrc, _
                                         ByRef intKtbnStrcSeqNo As Integer, _
                                         ByRef strOptionSymbol As String, _
                                         ByRef strMessageCd As String) As Boolean

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncCheckSelectOption = True

            If objKtbnStrc.strcSelection.strSeriesKataban.Trim = "SMG" Then
                If objKtbnStrc.strcSelection.strKeyKataban = "2" Then
                    Select Case Right(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, 1)
                        Case "0", "5"
                        Case Else
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W8320"
                            fncCheckSelectOption = False
                    End Select
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "X" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "Y" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 15 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "M" Then
                        If (objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "6" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "10" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "16") Then
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 30 Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0190"
                                fncCheckSelectOption = False
                            End If
                        Else
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 50 Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0190"
                                fncCheckSelectOption = False
                            End If
                        End If
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "F" Then
                        If (objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "6" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "10" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "16") Then
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 30 Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0190"
                                fncCheckSelectOption = False
                            End If
                        Else
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim > 50 Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0190"
                                fncCheckSelectOption = False
                            End If
                        End If
                    End If
                Else
                    Select Case Right(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, 1)
                        Case "0", "5"
                        Case Else
                            intKtbnStrcSeqNo = 6
                            strMessageCd = "W8320"
                            fncCheckSelectOption = False
                    End Select
                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "X" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "Y" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim > 15 Then
                            intKtbnStrcSeqNo = 6
                            strMessageCd = "W0190"
                            fncCheckSelectOption = False
                        End If
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "F" Then
                        If (objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "6" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "10" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "16") Then
                            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim > 30 Then
                                intKtbnStrcSeqNo = 6
                                strMessageCd = "W0190"
                                fncCheckSelectOption = False
                            End If
                        Else
                            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim > 50 Then
                                intKtbnStrcSeqNo = 6
                                strMessageCd = "W0190"
                                fncCheckSelectOption = False
                            End If
                        End If
                    End If
                    'RM1608009 Add Start K.Ohwaki 2016/08/29 クリーン仕様　ストローク制限
                    If (objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "P5" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "P51" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "P7" Or _
                        objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "P71") Then
                        If (objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "6" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "10" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(4).Trim = "16") Then
                            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim > 30 Then
                                intKtbnStrcSeqNo = 6
                                strMessageCd = "W0190"
                                fncCheckSelectOption = False
                            End If
                        Else
                            If objKtbnStrc.strcSelection.strOpSymbol(6).Trim > 50 Then
                                intKtbnStrcSeqNo = 6
                                strMessageCd = "W0190"
                                fncCheckSelectOption = False
                            End If
                        End If
                    End If
                    'RM1608009 Add End K.Ohwaki 2016/08/29

                End If
            End If

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "JSM2"
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "J"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    Next
                Case "JSM2-V"
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "J"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    Next
                Case "JSK2"
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "J"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    Next
                Case "JSK2-V"
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(8), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "J"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) < 25 Then
                                    intKtbnStrcSeqNo = 3
                                    strMessageCd = "W0190"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    Next
                Case "LN"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                        Case "LS"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "20"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 16 Then
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W0210"
                                        fncCheckSelectOption = False
                                    End If
                                Case "30"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 26 Then
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W0220"
                                        fncCheckSelectOption = False
                                    End If
                                Case "50"
                            End Select
                        Case "LDS"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                Case "20"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) = 4 Or CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 15 Then
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W0230"
                                        fncCheckSelectOption = False
                                    End If
                                Case "30"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) = 4 Or CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 25 Then
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W0240"
                                        fncCheckSelectOption = False
                                    End If
                                Case "50"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) = 4 Or CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim) > 45 Then
                                        intKtbnStrcSeqNo = 9
                                        strMessageCd = "W0250"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                    End Select
                Case "CAV2", "COVP2", "COVN2"
                    Dim bolOptionQ As Boolean = False

                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "Q"
                                bolOptionQ = True
                        End Select
                    Next
                    If bolOptionQ = False Then
                        If objKtbnStrc.strcSelection.strOpSymbol(6).Trim = "" Then
                            intKtbnStrcSeqNo = 6
                            strMessageCd = "W0260"
                            fncCheckSelectOption = False
                        End If
                    End If

                    Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                        Case "T0H", "T0V", "T5H", "T5V", "T8H", "T8V"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                Case "LB", "FA", "CA"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                        Case "H", "R"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "50"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 9 Then
                                                        intKtbnStrcSeqNo = 5
                                                        strMessageCd = "W0190"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "75", "100"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                                        intKtbnStrcSeqNo = 5
                                                        strMessageCd = "W0190"
                                                        fncCheckSelectOption = False
                                                    End If
                                            End Select
                                        Case "D"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "50"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 18 Then
                                                        intKtbnStrcSeqNo = 5
                                                        strMessageCd = "W0190"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "75", "100"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 19 Then
                                                        intKtbnStrcSeqNo = 5
                                                        strMessageCd = "W0190"
                                                        fncCheckSelectOption = False
                                                    End If
                                            End Select
                                        Case "T"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "50"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 35 Then
                                                        intKtbnStrcSeqNo = 5
                                                        strMessageCd = "W0190"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "75", "100"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 38 Then
                                                        intKtbnStrcSeqNo = 5
                                                        strMessageCd = "W0190"
                                                        fncCheckSelectOption = False
                                                    End If
                                            End Select
                                    End Select
                                Case "TC", "TF"
                                    'RM0911XXX 2009/11/10 Y.Miura ストローク制限変更
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < fncGetMinStroke_CAV2(objKtbnStrc) Then
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W0200"
                                        fncCheckSelectOption = False
                                    End If
                                    '        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    '            Case "N"
                                    '                Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                    '                    Case "R", "H", "D"
                                    '                        Select Case Right(Trim(objKtbnStrc.strcSelection.strOpSymbol(7).Trim), 1)
                                    '                            Case "V"
                                    '                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                                    Case "50"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 215 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "75"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 193 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "100"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 71 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                End Select
                                    '                            Case "H"
                                    '                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                                    Case "50"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 215 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "75"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 193 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "100"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 83 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                End Select
                                    '                        End Select
                                    '                    Case "T"
                                    '                        Select Case Right(Trim(objKtbnStrc.strcSelection.strOpSymbol(7).Trim), 1)
                                    '                            Case "V"
                                    '                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                                    Case "50"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 215 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "75"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 193 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "100"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 73 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                End Select
                                    '                            Case "H"
                                    '                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                                    Case "50"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 215 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "75"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 193 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "100"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 83 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                End Select
                                    '                        End Select
                                    '                End Select
                                    '            Case "B"
                                    '                Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                    '                    Case "R", "H", "D"
                                    '                        Select Case Right(Trim(objKtbnStrc.strcSelection.strOpSymbol(7).Trim), 1)
                                    '                            Case "V"
                                    '                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                                    Case "50"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 241 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "75"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 241 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "100"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 108 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                End Select
                                    '                            Case "H"
                                    '                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                                    Case "50"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 241 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "75"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 241 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "100"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 120 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                End Select
                                    '                        End Select
                                    '                    Case "T"
                                    '                        Select Case Right(Trim(objKtbnStrc.strcSelection.strOpSymbol(7).Trim), 1)
                                    '                            Case "V"
                                    '                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                                    Case "50"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 241 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "75"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 241 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "100"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 110 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                End Select
                                    '                            Case "H"
                                    '                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                                    Case "50"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 241 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "75"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 241 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                    Case "100"
                                    '                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 120 Then
                                    '                                            intKtbnStrcSeqNo = 5
                                    '                                            strMessageCd = "W0190"
                                    '                                            fncCheckSelectOption = False
                                    '                                        End If
                                    '                                End Select
                                    '                        End Select
                                    '                End Select
                                    '        End Select
                            End Select
                        Case "T2H", "T2V", "T3H", "T3V", "T2YH", "T2YV", "T3YH", "T3YV", "T1H", "T1V", _
                             "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", "T2YMV", "T3YMH", "T3YMV", "T2JH", "T2JV", "T2WH", "T3WH", "T2WV", "T3WV"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                Case "LB", "FA", "CA"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                        Case "H", "R"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "50", "75"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 5 Then
                                                        intKtbnStrcSeqNo = 5
                                                        strMessageCd = "W0190"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "100"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 6 Then
                                                        intKtbnStrcSeqNo = 5
                                                        strMessageCd = "W0190"
                                                        fncCheckSelectOption = False
                                                    End If
                                            End Select
                                        Case "D"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "50"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                                        intKtbnStrcSeqNo = 5
                                                        strMessageCd = "W0190"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "75"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 11 Then
                                                        intKtbnStrcSeqNo = 5
                                                        strMessageCd = "W0190"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "100"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 12 Then
                                                        intKtbnStrcSeqNo = 5
                                                        strMessageCd = "W0190"
                                                        fncCheckSelectOption = False
                                                    End If
                                            End Select
                                        Case "T"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "50"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 20 Then
                                                        intKtbnStrcSeqNo = 5
                                                        strMessageCd = "W0190"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "75"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 21 Then
                                                        intKtbnStrcSeqNo = 5
                                                        strMessageCd = "W0190"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "100"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 23 Then
                                                        intKtbnStrcSeqNo = 5
                                                        strMessageCd = "W0190"
                                                        fncCheckSelectOption = False
                                                    End If
                                            End Select
                                    End Select
                                Case "TC", "TF"
                                    'RM0911XXX 2009/11/10 Y.Miura ストローク制限変更
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < fncGetMinStroke_CAV2(objKtbnStrc) Then
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W0200"
                                        fncCheckSelectOption = False
                                    End If
                                    'Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    '    Case "N"
                                    '        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                    '            Case "R", "H", "D"
                                    '                Select Case Right(Trim(objKtbnStrc.strcSelection.strOpSymbol(7).Trim), 1)
                                    '                    Case "V"
                                    '                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                            Case "50"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 46 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "75"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 24 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "100"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 54 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                        End Select
                                    '                    Case "H"
                                    '                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                            Case "50"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 76 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "75"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 54 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "100"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 84 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                        End Select
                                    '                End Select
                                    '            Case "T"
                                    '                Select Case Right(Trim(objKtbnStrc.strcSelection.strOpSymbol(7).Trim), 1)
                                    '                    Case "V"
                                    '                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                            Case "50"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 47 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "75"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 26 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "100"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 58 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                        End Select
                                    '                    Case "H"
                                    '                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                            Case "50"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 76 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "75"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 54 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "100"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 84 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                        End Select
                                    '                End Select
                                    '        End Select
                                    '    Case "B"
                                    '        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                    '            Case "R", "H", "D"
                                    '                Select Case Right(Trim(objKtbnStrc.strcSelection.strOpSymbol(7).Trim), 1)
                                    '                    Case "V"
                                    '                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                            Case "50"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 72 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "75"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 72 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "100"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 91 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                        End Select
                                    '                    Case "H"
                                    '                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                            Case "50"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 102 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "75"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 102 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "100"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 121 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                        End Select
                                    '                End Select
                                    '            Case "T"
                                    '                Select Case Right(Trim(objKtbnStrc.strcSelection.strOpSymbol(7).Trim), 1)
                                    '                    Case "V"
                                    '                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                            Case "50"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 73 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "75"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 74 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "100"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 95 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                        End Select
                                    '                    Case "H"
                                    '                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    '                            Case "50"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 102 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "75"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 102 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                            Case "100"
                                    '                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 121 Then
                                    '                                    intKtbnStrcSeqNo = 5
                                    '                                    strMessageCd = "W0190"
                                    '                                    fncCheckSelectOption = False
                                    '                                End If
                                    '                        End Select
                                    '                End Select
                                    '        End Select
                                    'End Select
                            End Select
                    End Select
                Case "LCS", "LCS-Q"
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case ""
                            Case "A1", "A2", "A5", "A6"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "6", "8"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = 10 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0270"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "12", "16", "20", "25"
                                        'RMXXXXXXX 2009/09/11 Y.Miura 不具合修正
                                        'If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = 10 Or _
                                        '    intKtbnStrcSeqNo = 3 Then
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = 10 Or _
                                          CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = 20 Then
                                            intKtbnStrcSeqNo = 3
                                            strMessageCd = "W0270"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select
                    Next
                Case "STG-B", "STG-M", "STG-K"
                    Select Case Right(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, 1)
                        Case "0", "5"
                        Case Else
                            intKtbnStrcSeqNo = 4
                            strMessageCd = "W8320"
                            fncCheckSelectOption = False
                    End Select

                    If objKtbnStrc.strcSelection.strOpSymbol(6).Trim <> "" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                            Case "T"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) < 25 Then
                                    intKtbnStrcSeqNo = 4
                                    strMessageCd = "W0200"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "C" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            Case "16"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "25", "50", "75", "100", "125", _
                                         "150", "175", "200", "250"
                                    Case Else
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W0280"
                                        fncCheckSelectOption = False
                                End Select
                            Case Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                    Case "25", "50", "75", "100", "125", _
                                         "150", "175", "200", "250", "300", _
                                         "350", "400"
                                    Case Else
                                        intKtbnStrcSeqNo = 4
                                        strMessageCd = "W0280"
                                        fncCheckSelectOption = False
                                End Select
                        End Select
                    End If
                    '2013/06/27 グローバル機種対応(SCW)
                    'RM1712042_SCWP2,SCWT2追加
                Case "SCW", "SCWP2", "SCWT2"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "TA", "TB"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 1900 Then  'RM1806044_ストローク対応
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0190"
                                fncCheckSelectOption = False
                            End If
                    End Select
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                        Case "R", "H"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncCheckSelectOption = False
                            End If
                        Case "D"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncCheckSelectOption = False
                            End If
                        Case "T"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 30 Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0200"
                                fncCheckSelectOption = False
                            End If
                    End Select
                    If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(11).Trim, "B1") <> 0 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "CB" Then
                        Else
                            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(11).Trim, "Y") = 0 Then
                                intKtbnStrcSeqNo = 11
                                strMessageCd = "W0290"
                                fncCheckSelectOption = False
                            End If
                        End If
                    End If
                    If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(11).Trim, "B2") <> 0 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "CA" Then
                        Else
                            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(11).Trim, "I") = 0 Then
                                intKtbnStrcSeqNo = 11
                                strMessageCd = "W0300"
                                fncCheckSelectOption = False
                            End If
                        End If
                    End If
                    If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(11).Trim, "B3") <> 0 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "CB" Then
                        Else
                            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(11).Trim, "Y") = 0 Then
                                intKtbnStrcSeqNo = 11
                                strMessageCd = "W0310"
                                fncCheckSelectOption = False
                            End If
                        End If
                    End If
                Case "SCG", "SCG-D", "SCG-G", "SCG-G2", "SCG-G3", _
                     "SCG-G4", "SCG-M", "SCG-O", "SCG-Q", "SCG-U", "SCG-G1", "SCG-G1L2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "F"
                            '食品製造工程向け商品
                            If objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("B1") >= 0 Then
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "CB" Then
                                Else
                                    If objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("Y") < 0 Then
                                        intKtbnStrcSeqNo = 11
                                        strMessageCd = "W0290"
                                        fncCheckSelectOption = False
                                    End If
                                End If
                            End If

                            If objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("B2") >= 0 Then
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "CA" Then
                                Else
                                    If objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("I") < 0 Then
                                        intKtbnStrcSeqNo = 11
                                        strMessageCd = "W0300"
                                        fncCheckSelectOption = False
                                    End If
                                End If
                            End If

                            If objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("B3") >= 0 Then
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "CB" Then
                                Else
                                    If objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("Y") < 0 Then
                                        intKtbnStrcSeqNo = 11
                                        strMessageCd = "W0310"
                                        fncCheckSelectOption = False
                                    End If
                                End If
                            End If
                        Case Else
                            If objKtbnStrc.strcSelection.strOpSymbol(11).IndexOf("B1") >= 0 Then
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "CB" Then
                                Else
                                    If objKtbnStrc.strcSelection.strOpSymbol(11).IndexOf("Y") < 0 Then
                                        intKtbnStrcSeqNo = 11
                                        strMessageCd = "W0290"
                                        fncCheckSelectOption = False
                                    End If
                                End If
                            End If

                            If objKtbnStrc.strcSelection.strOpSymbol(11).IndexOf("B2") >= 0 Then
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "CA" Then
                                Else
                                    If objKtbnStrc.strcSelection.strOpSymbol(11).IndexOf("I") < 0 Then
                                        intKtbnStrcSeqNo = 11
                                        strMessageCd = "W0300"
                                        fncCheckSelectOption = False
                                    End If
                                End If
                            End If

                            If objKtbnStrc.strcSelection.strOpSymbol(11).IndexOf("B3") >= 0 Then
                                If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "CB" Then
                                Else
                                    If objKtbnStrc.strcSelection.strOpSymbol(11).IndexOf("Y") < 0 Then
                                        intKtbnStrcSeqNo = 11
                                        strMessageCd = "W0310"
                                        fncCheckSelectOption = False
                                    End If
                                End If
                            End If
                    End Select

                    'RM0911XXX 2009/11/10 Y.Miura ストローク制限変更
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < fncGetMinStroke_SCG(objKtbnStrc) Then
                        intKtbnStrcSeqNo = 5
                        strMessageCd = "W0200"
                        fncCheckSelectOption = False
                    End If
                    'If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                    '    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                    '        Case "00", "LB", "FA", "FB", "CA", "CB"
                    '            Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                Case "H", "R"
                    '                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                    '                        intKtbnStrcSeqNo = 5
                    '                        strMessageCd = "W0200"
                    '                        fncCheckSelectOption = False
                    '                    End If
                    '                Case "D"
                    '                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                    '                        intKtbnStrcSeqNo = 5
                    '                        strMessageCd = "W0200"
                    '                        fncCheckSelectOption = False
                    '                    End If
                    '                Case "T"
                    '                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 30 Then
                    '                        intKtbnStrcSeqNo = 5
                    '                        strMessageCd = "W0200"
                    '                        fncCheckSelectOption = False
                    '                    End If
                    '                Case Else
                    '                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 35 Then
                    '                        intKtbnStrcSeqNo = 5
                    '                        strMessageCd = "W0200"
                    '                        fncCheckSelectOption = False
                    '                    End If
                    '            End Select
                    '        Case "TC"
                    '            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    '                Case "32"
                    '                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                        Case "H", "R", "D"
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 63 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                        Case Else
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 93 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                    End Select
                    '                Case "40", "50"
                    '                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                        Case "H", "R", "D"
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 68 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                        Case Else
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 98 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                    End Select
                    '                Case "63"
                    '                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                        Case "H", "R", "D"
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 74 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                        Case Else
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 98 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                    End Select
                    '                Case "80"
                    '                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                        Case "H", "R", "D"
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 86 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                        Case Else
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 101 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                    End Select
                    '                Case "100"
                    '                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                        Case "H", "R", "D"
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 92 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                        Case Else
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 107 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                    End Select
                    '            End Select
                    '        Case "TA"
                    '            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    '                Case "32"
                    '                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                        Case "H"
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 37 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                    End Select
                    '                Case "40", "50"
                    '                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                        Case "H"
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 42 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                    End Select
                    '                Case "63"
                    '                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                        Case "H"
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 48 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                    End Select
                    '                Case "80"
                    '                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                        Case "H"
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 54 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                    End Select
                    '                Case "100"
                    '                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                        Case "H"
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 60 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                    End Select
                    '            End Select
                    '        Case "TB"
                    '            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                    '                Case "32"
                    '                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                        Case "R"
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 37 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                    End Select
                    '                Case "40", "50"
                    '                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                        Case "R"
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 42 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                    End Select
                    '                Case "63"
                    '                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                        Case "R"
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 48 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                    End Select
                    '                Case "80"
                    '                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                        Case "R"
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 54 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                    End Select
                    '                Case "100"
                    '                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                    '                        Case "R"
                    '                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 60 Then
                    '                                intKtbnStrcSeqNo = 5
                    '                                strMessageCd = "W0200"
                    '                                fncCheckSelectOption = False
                    '                            End If
                    '                    End Select
                    '            End Select
                    '    End Select
                    'End If
                Case "CMK2"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("R") >= 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("SR") < 0 Then
                                Select Case CInt(objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                                    Case 1 To 25
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(11).Trim
                                            Case "25"
                                            Case Else
                                                intKtbnStrcSeqNo = 11
                                                strMessageCd = "W0320"
                                                fncCheckSelectOption = False
                                        End Select
                                End Select
                            End If
                        End If
                    End If
                Case "SCM"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("R") >= 0 Then
                            Select Case CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                                Case 1 To 25
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                        Case "25"
                                        Case Else
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0320"
                                            fncCheckSelectOption = False
                                    End Select
                            End Select
                        End If
                    End If
                Case "SCA2"
                    If objKtbnStrc.strcSelection.strKeyKataban.Trim = "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("R") >= 0 Then
                            Select Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                                Case 1 To 25
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                        Case "25"
                                        Case Else
                                            intKtbnStrcSeqNo = 9
                                            strMessageCd = "W0320"
                                            fncCheckSelectOption = False
                                    End Select
                                Case 26 To 50
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                        Case "25", "50"
                                        Case Else
                                            intKtbnStrcSeqNo = 9
                                            strMessageCd = "W0320"
                                            fncCheckSelectOption = False
                                    End Select
                                Case 51 To 75
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                        Case "25", "50", "75"
                                        Case Else
                                            intKtbnStrcSeqNo = 9
                                            strMessageCd = "W0320"
                                            fncCheckSelectOption = False
                                    End Select
                            End Select
                        End If
                    End If

                    'SCA2-Vのみチェック
                    If objKtbnStrc.strcSelection.strKeyKataban = "V" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Or _
                           objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("R") >= 0 Then
                            Select Case CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                                Case 1 To 25
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                        Case "25"
                                        Case Else
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0320"
                                            fncCheckSelectOption = False
                                    End Select
                                Case 26 To 50
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                        Case "25", "50"
                                        Case Else
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0320"
                                            fncCheckSelectOption = False
                                    End Select
                                Case 51 To 75
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                                        Case "25", "50", "75"
                                        Case Else
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0320"
                                            fncCheckSelectOption = False
                                    End Select
                            End Select
                        End If
                    End If
                    '↓RM1302XXX 2013/02/04 Y.Tachi
                Case "SCS", "SCS2"
                    Dim intFlg As Integer = 0
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        '2011/10/24 MOD RM1110032(11月VerUP:二次電池) START--->
                        Case "", "D", "2", "4", "F", "G"
                            'Case "", "D"
                            '2011/10/24 MOD RM1110032(11月VerUP:二次電池) <---END
                            If IsNumeric(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) = False Then
                                '数値以外は不可
                                intFlg = 1
                            End If

                            For intLoopCnt = 1 To Len(objKtbnStrc.strcSelection.strOpSymbol(12).Trim)
                                If Mid(objKtbnStrc.strcSelection.strOpSymbol(12).Trim, intLoopCnt, 1) = CdCst.Sign.Dot Then
                                    intFlg = 1
                                    Exit For
                                End If
                            Next

                            If intFlg = 1 Then
                                intKtbnStrcSeqNo = 12
                                strMessageCd = "W0330"
                                fncCheckSelectOption = False
                            End If
                        Case "B"
                            If IsNumeric(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) = False Then
                                '数値以外は不可
                                intFlg = 1
                            End If

                            For intLoopCnt = 1 To Len(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                                If Mid(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, intLoopCnt, 1) = CdCst.Sign.Dot Then
                                    intFlg = 1
                                    Exit For
                                End If
                            Next

                            If intFlg = 1 Then
                                intKtbnStrcSeqNo = 6
                                strMessageCd = "W0330"
                                fncCheckSelectOption = False
                            End If
                    End Select

                    '二段形の時、S1とS2の大小関係をチェックする
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) >= CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) Then
                        Else
                            intKtbnStrcSeqNo = 12
                            strMessageCd = "W0610"
                            fncCheckSelectOption = False
                            Exit Try
                        End If
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                            Case "125", "140", "160"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 800 Then
                                    intKtbnStrcSeqNo = 12
                                    strMessageCd = "W0200"
                                    fncCheckSelectOption = False
                                End If
                            Case "180"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 900 Then
                                    intKtbnStrcSeqNo = 12
                                    strMessageCd = "W0200"
                                    fncCheckSelectOption = False
                                End If
                            Case "200"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 1000 Then
                                    intKtbnStrcSeqNo = 12
                                    strMessageCd = "W0200"
                                    fncCheckSelectOption = False
                                End If
                            Case "250"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 1200 Then
                                    intKtbnStrcSeqNo = 12
                                    strMessageCd = "W0200"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W") >= 0 Then
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 200 Then
                            intKtbnStrcSeqNo = 12
                            strMessageCd = "W0200"
                            fncCheckSelectOption = False
                        End If
                    End If

                    Select Case True
                        Case 1 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) And CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) <= 25
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                                Case "", "25"
                                Case Else
                                    intKtbnStrcSeqNo = 13
                                    strMessageCd = "W0320"
                                    fncCheckSelectOption = False
                            End Select
                        Case 26 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) And CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) <= 50
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                                Case "", "25", "50"
                                Case Else
                                    intKtbnStrcSeqNo = 13
                                    strMessageCd = "W0320"
                                    fncCheckSelectOption = False
                            End Select
                        Case 51 <= CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) And CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) <= 75
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(13).Trim
                                Case "", "25", "50", "75"
                                Case Else
                                    intKtbnStrcSeqNo = 13
                                    strMessageCd = "W0320"
                                    fncCheckSelectOption = False
                            End Select
                    End Select

                    '2011/10/24 MOD RM1110032(11月VerUP:二次電池) START--->
                    '付属品の位置を取得
                    Dim intFuzoku As Integer
                    If objKtbnStrc.strcSelection.strKeyKataban = "2" Or objKtbnStrc.strcSelection.strSeriesKataban = "SCS2" Then
                        If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                            intFuzoku = 20
                        Else
                            intFuzoku = 19
                        End If
                    Else
                        intFuzoku = 18
                    End If

                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intFuzoku), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "B1"
                                If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "CB" And _
                                   objKtbnStrc.strcSelection.strOpSymbol(intFuzoku).IndexOf("Y") < 0 Then
                                    intKtbnStrcSeqNo = intFuzoku
                                    strMessageCd = "W0290"
                                    fncCheckSelectOption = False
                                End If
                            Case "B2"
                                If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "CA" And _
                                   objKtbnStrc.strcSelection.strOpSymbol(intFuzoku).IndexOf("I") < 0 Then
                                    intKtbnStrcSeqNo = intFuzoku
                                    strMessageCd = "W0300"
                                    fncCheckSelectOption = False
                                End If
                        End Select
                    Next
                    'strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(18), CdCst.Sign.Delimiter.Comma)
                    'For intLoopCnt = 0 To strOpArray.Length - 1
                    '    Select Case strOpArray(intLoopCnt).Trim
                    '        Case "B1"
                    '            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "CB" And _
                    '               objKtbnStrc.strcSelection.strOpSymbol(18).IndexOf("Y") < 0 Then
                    '                intKtbnStrcSeqNo = 18
                    '                strMessageCd = "W0290"
                    '                fncCheckSelectOption = False
                    '            End If
                    '        Case "B2"
                    '            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "CA" And _
                    '               objKtbnStrc.strcSelection.strOpSymbol(18).IndexOf("I") < 0 Then
                    '                intKtbnStrcSeqNo = 18
                    '                strMessageCd = "W0300"
                    '                fncCheckSelectOption = False
                    '            End If
                    '    End Select
                    'Next
                    '2011/10/24 MOD RM1110032(11月VerUP:二次電池) <---END

                    If objKtbnStrc.strcSelection.strKeyKataban = "B" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(18).IndexOf("L") < 0 Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                Case "TD", "TE", "TF"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "125"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 30 Then
                                                intKtbnStrcSeqNo = 6
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "140"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 32 Then
                                                intKtbnStrcSeqNo = 6
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "160"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 34 Then
                                                intKtbnStrcSeqNo = 6
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "180"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 35 Then
                                                intKtbnStrcSeqNo = 6
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "200"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 37 Then
                                                intKtbnStrcSeqNo = 6
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "250"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 39 Then
                                                intKtbnStrcSeqNo = 6
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "TA", "TB", "TC"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "125"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 23 Then
                                                intKtbnStrcSeqNo = 6
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "140"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 25 Then
                                                intKtbnStrcSeqNo = 6
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "160"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 27 Then
                                                intKtbnStrcSeqNo = 6
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "180"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 28 Then
                                                intKtbnStrcSeqNo = 6
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "200"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 28 Then
                                                intKtbnStrcSeqNo = 6
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "250"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 28 Then
                                                intKtbnStrcSeqNo = 6
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select

                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "00", "LB", "FA", "FB", "CA", "CB"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "TD", "TE", "TF"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            Case "125"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 30 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "140"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 32 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "160"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 34 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "180"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 35 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "200"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 37 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "250"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 39 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "TA", "TB", "TC"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            Case "125"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 23 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "140"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 25 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "160"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 27 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "180"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 28 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "200"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 28 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "250"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 28 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select
                            End If

                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("H") >= 0 Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "00", "LB", "FA", "FB", "CA", "CB"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 20 Then
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "TD", "TE", "TF"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            Case "125"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 30 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "140"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 32 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "160"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 34 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "180"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 35 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "200"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 37 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "250"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 39 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "TA", "TB", "TC"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            Case "125"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 23 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "140"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 25 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "160"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 27 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "180"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 28 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "200"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 28 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "250"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 28 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select
                            End If
                        Else
                            If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "" Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "00", "LB", "FA", "FB", "CA", "CB"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 20 Then
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "TC", "TF"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            Case "125"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 120 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "140"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 125 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "160"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 130 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "180"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 135 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "200"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 140 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "250"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 150 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "TA", "TD", "TB", "TE"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            Case "125"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 70 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "140"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 75 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "160"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 80 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "180"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 85 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "200"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 90 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "250"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 100 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select

                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Then
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                        Case "00", "LB", "FA", "FB", "CA", "CB"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 25 Then
                                                intKtbnStrcSeqNo = 6
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "TC", "TF"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "125"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 120 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "140"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 125 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "160"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 130 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "180"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 135 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "200"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 140 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "250"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 150 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                            End Select
                                        Case "TA", "TD", "TB", "TE"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "125"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 70 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "140"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 75 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "160"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 80 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "180"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 85 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "200"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 90 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "250"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 100 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                            End Select
                                    End Select
                                End If
                            Else
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "00", "LB", "FA", "FB", "CA", "CB"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "H", "R", "D"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 20 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "T"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 40 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "4"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 55 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "TC", "TF"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            Case "125"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 120 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "140"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 125 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "160"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 130 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "180"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 135 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "200"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 140 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "TA", "TD", "TB", "TE"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            Case "125"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 70 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "140"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 75 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "160"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 80 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "180"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 85 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "200"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 90 Then
                                                    intKtbnStrcSeqNo = 6
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select

                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Then
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                        Case "00", "LB", "FA", "FB", "CA", "CB"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                                Case "H", "R", "D"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 25 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "T"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 40 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "4"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 55 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                            End Select
                                        Case "TC", "TF"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "125"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 120 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "140"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 125 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "160"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 130 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "180"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 135 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "200"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 140 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                            End Select
                                        Case "TA", "TD", "TB", "TE"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "125"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 70 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "140"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 75 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "160"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 80 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "180"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 85 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "200"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < 90 Then
                                                        intKtbnStrcSeqNo = 6
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                            End Select
                                    End Select
                                End If
                            End If
                        End If
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("L") < 0 Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                            Case "TD", "TE", "TF"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    Case "125"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 30 Then
                                            intKtbnStrcSeqNo = 12
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "140"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 32 Then
                                            intKtbnStrcSeqNo = 12
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "160"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 34 Then
                                            intKtbnStrcSeqNo = 12
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "180"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 35 Then
                                            intKtbnStrcSeqNo = 12
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "200"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 37 Then
                                            intKtbnStrcSeqNo = 12
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "250"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 39 Then
                                            intKtbnStrcSeqNo = 12
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "TA", "TB", "TC"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                    Case "125"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 23 Then
                                            intKtbnStrcSeqNo = 12
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "140"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 12
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "160"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 27 Then
                                            intKtbnStrcSeqNo = 12
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "180"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 28 Then
                                            intKtbnStrcSeqNo = 12
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "200"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 28 Then
                                            intKtbnStrcSeqNo = 12
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "250"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 28 Then
                                            intKtbnStrcSeqNo = 12
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                        End Select

                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                Case "00", "LB", "FA", "FB", "CA", "CB"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 25 Then
                                        intKtbnStrcSeqNo = 12
                                        strMessageCd = "W0200"
                                        fncCheckSelectOption = False
                                    End If
                                Case "TD", "TE", "TF"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "125"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 30 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "140"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 32 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "160"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 34 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "180"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 35 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "200"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 37 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "250"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 39 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "TA", "TB", "TC"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "125"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 23 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "140"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 25 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "160"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 27 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "180"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 28 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "200"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 28 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "250"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 28 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("H") >= 0 Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                Case "00", "LB", "FA", "FB", "CA", "CB"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 20 Then
                                        intKtbnStrcSeqNo = 12
                                        strMessageCd = "W0200"
                                        fncCheckSelectOption = False
                                    End If
                                Case "TD", "TE", "TF"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "125"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 30 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "140"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 32 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "160"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 34 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "180"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 35 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "200"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 37 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "250"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 39 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "TA", "TB", "TC"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "125"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 23 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "140"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 25 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "160"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 27 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "180"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 28 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "200"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 28 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "250"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 28 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select
                        End If
                    Else
                        If objKtbnStrc.strcSelection.strOpSymbol(14).Trim = "" Then
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                Case "00", "LB", "FA", "FB", "CA", "CB"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 20 Then
                                        intKtbnStrcSeqNo = 12
                                        strMessageCd = "W0200"
                                        fncCheckSelectOption = False
                                    End If
                                Case "TC", "TF"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "125"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 120 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "140"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 125 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "160"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 130 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "180"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "200"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 140 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "TA", "TD", "TB", "TE"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "125"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 70 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "140"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 75 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "160"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 80 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "180"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 85 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "200"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select

                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "00", "LB", "FA", "FB", "CA", "CB"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 12
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "TC", "TF"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            Case "125"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 120 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "140"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 125 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "160"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 130 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "180"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 135 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "200"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 140 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "TA", "TD", "TB", "TE"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            Case "125"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 70 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "140"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 75 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "160"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 80 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "180"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 85 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "200"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 90 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select
                            End If
                        Else
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                Case "00", "LB", "FA", "FB", "CA", "CB"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                                        Case "H", "R", "D"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 20 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "T"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 40 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "4"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 55 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                                Case "TC", "TF"
                                    If InStr(1, objKtbnStrc.strcSelection.strOtherOption, "AQ") <> 0 Then
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            Case "125"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 30 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "140"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 32 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "160"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 34 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "180"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 35 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "200"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 37 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "250"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 39 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Else
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            Case "125"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 120 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "140"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 125 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "160"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 130 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "180"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 135 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "200"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 140 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "250"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 150 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    End If
                                Case "TA", "TD", "TB", "TE"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                        Case "125"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 70 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "140"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 75 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "160"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 80 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "180"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 85 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "200"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 90 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                        Case "250"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 100 Then
                                                intKtbnStrcSeqNo = 12
                                                strMessageCd = "W0200"
                                                fncCheckSelectOption = False
                                            End If
                                    End Select
                            End Select

                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Then
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "00", "LB", "FA", "FB", "CA", "CB"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                                            Case "H", "R", "D"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 25 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "T"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 40 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "4"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 55 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "TC", "TF"
                                        If InStr(1, objKtbnStrc.strcSelection.strOtherOption, "AQ") <> 0 Then
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "125"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 30 Then
                                                        intKtbnStrcSeqNo = 12
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "140"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 32 Then
                                                        intKtbnStrcSeqNo = 12
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "160"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 34 Then
                                                        intKtbnStrcSeqNo = 12
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "180"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 35 Then
                                                        intKtbnStrcSeqNo = 12
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "200"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 37 Then
                                                        intKtbnStrcSeqNo = 12
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "250"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 39 Then
                                                        intKtbnStrcSeqNo = 12
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                            End Select
                                        Else
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                                Case "125"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 120 Then
                                                        intKtbnStrcSeqNo = 12
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "140"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 125 Then
                                                        intKtbnStrcSeqNo = 12
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "160"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 130 Then
                                                        intKtbnStrcSeqNo = 12
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "180"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 135 Then
                                                        intKtbnStrcSeqNo = 12
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                                Case "200"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 140 Then
                                                        intKtbnStrcSeqNo = 12
                                                        strMessageCd = "W0200"
                                                        fncCheckSelectOption = False
                                                    End If
                                            End Select
                                        End If
                                    Case "TA", "TD", "TB", "TE"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                                            Case "125"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 70 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "140"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 75 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "160"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 80 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "180"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 85 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "200"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) < 90 Then
                                                    intKtbnStrcSeqNo = 12
                                                    strMessageCd = "W0200"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select
                            End If
                        End If
                    End If

                    '↓RM1401080 2014/01/27
                    If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(1).Trim, "W") <> 0 Then
                        If Val(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) >= Val(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) Then
                        Else
                            intKtbnStrcSeqNo = 12
                            strMessageCd = "W0610"
                            fncCheckSelectOption = False
                        End If
                    End If

                    'If Len(Trim(Rod_FulPartsNo)) <> 0 Then
                    '    Select Case Trim(SrsPartsNo)
                    '        Case "SCS"
                    '            ' オプション「A2」選択時はロッド先端は選択不可
                    '            If InStr(1, ItemCode(17), "A2") <> 0 Then
                    '                If InStr(1, Rod_FulPartsNo, "N13") <> 0 Or _
                    '                   InStr(1, Rod_FulPartsNo, "N15") <> 0 Then
                    '                Else
                    '                    intKtbnStrcSeqNo = 17
                    '                    strMessageCd = "W0340"
                    '                    fncCheckSelectOption = False
                    '                End If
                    '            End If

                    '            ' 付属品「I」「Y」選択時はロッド先端は選択不可
                    '            If InStr(1, ItemCode(18), "I") <> 0 Or _
                    '               InStr(1, ItemCode(18), "Y") <> 0 Then
                    '                If InStr(1, Rod_FulPartsNo, "N13") <> 0 Or _
                    '                   InStr(1, Rod_FulPartsNo, "N15") <> 0 Then
                    '                Else
                    '                    intKtbnStrcSeqNo = 18
                    '                    strMessageCd = "W0350"
                    '                    fncCheckSelectOption = False
                    '                End If
                    '            End If
                    '        Case "SCS      B"
                    '            ' バリエーション「B」選択時はロッド先端は選択不可
                    '            If InStr(1, ItemCode(1), "B") <> 0 Then
                    '                If Len(Trim(Rod_FulPartsNo)) <> 0 Then
                    '                    intKtbnStrcSeqNo = 18
                    '                    strMessageCd = "W0360"
                    '                    fncCheckSelectOption = False
                    '                End If
                    '            End If

                    '            ' オプション「A2」選択時はロッド先端は選択不可
                    '            If InStr(1, ItemCode(17), "A2") <> 0 Then
                    '                If InStr(1, Rod_FulPartsNo, "N13") <> 0 Or _
                    '                   InStr(1, Rod_FulPartsNo, "N15") <> 0 Then
                    '                Else
                    '                    intKtbnStrcSeqNo = 17
                    '                    strMessageCd = "W0370"
                    '                    fncCheckSelectOption = False
                    '                End If
                    '            End If

                    '            ' 付属品「I」「Y」選択時はロッド先端は選択不可
                    '            If InStr(1, ItemCode(18), "I") <> 0 Or _
                    '               InStr(1, ItemCode(18), "Y") <> 0 Then
                    '                If InStr(1, Rod_FulPartsNo, "N13") <> 0 Or _
                    '                   InStr(1, Rod_FulPartsNo, "N15") <> 0 Then
                    '                Else
                    '                    intKtbnStrcSeqNo = 18
                    '                    strMessageCd = "W0350"
                    '                    fncCheckSelectOption = False
                    '                End If
                    '            End If
                    '        Case "SCS      D"
                    '            ' 付属品「IY」選択時はロッド先端は選択不可
                    '            If InStr(1, ItemCode(18), "IY") <> 0 Then
                    '                If InStr(1, Rod_FulPartsNo, "N13-N11") <> 0 Then
                    '                    intKtbnStrcSeqNo = 18
                    '                    strMessageCd = "W0380"
                    '                    fncCheckSelectOption = False
                    '                End If
                    '            End If
                    '    End Select
                    'End If

                    '2011/01/31 MOD RM1101046(2月VerUP：オプション外特注処理追加) START--->
                    ' オプション外
                    If Len(Trim(objKtbnStrc.strcSelection.strOtherOption)) <> 0 Then
                        ' クッションニードル位置指定
                        If Left(Trim(objKtbnStrc.strcSelection.strOtherOption), 1) = "R" Then
                            If Trim(objKtbnStrc.strcSelection.strOpSymbol(5)) <> "B" Then
                                intKtbnStrcSeqNo = 5
                                strMessageCd = "W0390"
                                fncCheckSelectOption = False
                            End If

                            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(17), "S") <> 0 Or _
                               InStr(1, objKtbnStrc.strcSelection.strOpSymbol(17), "T") <> 0 Then
                                intKtbnStrcSeqNo = 17
                                strMessageCd = "W0400"
                                fncCheckSelectOption = False
                            End If
                        End If

                        ' ポート２箇所指定
                        If InStr(1, objKtbnStrc.strcSelection.strOtherOption, "E") <> 0 Then
                            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(17), "S") <> 0 Then
                                intKtbnStrcSeqNo = 17
                                strMessageCd = "W0410"
                                fncCheckSelectOption = False
                            End If
                        End If

                        ' 支持金具90°回転
                        If InStr(1, objKtbnStrc.strcSelection.strOtherOption, "K1") <> 0 Then
                            If Trim(objKtbnStrc.strcSelection.strOpSymbol(2)) = "00" Then
                                intKtbnStrcSeqNo = 2
                                strMessageCd = "W0430"
                                fncCheckSelectOption = False
                            End If
                        End If

                        ' 支持金具180°回転
                        If InStr(1, objKtbnStrc.strcSelection.strOtherOption, "K2") <> 0 Then
                            If Trim(objKtbnStrc.strcSelection.strOpSymbol(2)) <> "LB" Then
                                intKtbnStrcSeqNo = 2
                                strMessageCd = "W0440"
                                fncCheckSelectOption = False
                            End If
                        End If

                        ' 支持金具270°回転
                        If InStr(1, objKtbnStrc.strcSelection.strOtherOption, "K3") <> 0 Then
                            If Trim(objKtbnStrc.strcSelection.strOpSymbol(2)) <> "LB" Then
                                intKtbnStrcSeqNo = 2
                                strMessageCd = "W0450"
                                fncCheckSelectOption = False
                            End If
                        End If

                        ' トラニオン位置
                        If InStr(1, objKtbnStrc.strcSelection.strOtherOption, "AQ") <> 0 Then
                            If Trim(objKtbnStrc.strcSelection.strOpSymbol(2)) <> "TC" And _
                               Trim(objKtbnStrc.strcSelection.strOpSymbol(2)) <> "TF" Then
                                intKtbnStrcSeqNo = 2
                                strMessageCd = "W0460"
                                fncCheckSelectOption = False
                            End If
                        End If

                        ' P5
                        'RM1210067 2013/02/01 Y.Tachi ローカル版との差異修正
                        '↓RM1310004 2013/10/01 追加
                        If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "SCS2" Then
                            If InStr(1, objKtbnStrc.strcSelection.strOtherOption, "P5") <> 0 Then
                                If Trim(objKtbnStrc.strcSelection.strOpSymbol(2)) <> "CB" AndAlso _
                                   InStr(Trim(objKtbnStrc.strcSelection.strOpSymbol(19)), "Y") = 0 Then
                                    intKtbnStrcSeqNo = 19
                                    strMessageCd = "W0470"
                                    fncCheckSelectOption = False
                                End If
                            End If
                        Else
                            If InStr(1, objKtbnStrc.strcSelection.strOtherOption, "P5") <> 0 Then
                                If Trim(objKtbnStrc.strcSelection.strOpSymbol(2)) <> "CB" AndAlso _
                                   InStr(Trim(objKtbnStrc.strcSelection.strOpSymbol(18)), "Y") = 0 Then
                                    intKtbnStrcSeqNo = 18
                                    strMessageCd = "W0470"
                                    fncCheckSelectOption = False
                                End If
                            End If
                        End If

                        ' P7
                        '↓RM1310004 2013/10/01 追加
                        If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "SCS2" Then
                            If InStr(1, objKtbnStrc.strcSelection.strOtherOption, "P7") <> 0 Then
                                If InStr(Trim(objKtbnStrc.strcSelection.strOpSymbol(19)), "I") = 0 AndAlso _
                                   InStr(Trim(objKtbnStrc.strcSelection.strOpSymbol(19)), "Y") = 0 Then
                                    intKtbnStrcSeqNo = 19
                                    strMessageCd = "W0480"
                                    fncCheckSelectOption = False
                                End If
                            End If
                        Else
                            If InStr(1, objKtbnStrc.strcSelection.strOtherOption, "P7") <> 0 Then
                                If InStr(Trim(objKtbnStrc.strcSelection.strOpSymbol(18)), "I") = 0 AndAlso _
                                   InStr(Trim(objKtbnStrc.strcSelection.strOpSymbol(18)), "Y") = 0 Then
                                    intKtbnStrcSeqNo = 18
                                    strMessageCd = "W0480"
                                    fncCheckSelectOption = False
                                End If
                            End If
                        End If

                        ' P8
                        '↓RM1310004 2013/10/01 追加
                        If Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4) = "SCS2" Then
                            If InStr(1, objKtbnStrc.strcSelection.strOtherOption, "P8") <> 0 Then
                                If InStr(Trim(objKtbnStrc.strcSelection.strOpSymbol(19)), "Y") = 0 Then
                                    intKtbnStrcSeqNo = 19
                                    strMessageCd = "W0490"
                                    fncCheckSelectOption = False
                                End If
                            End If
                        Else
                            If InStr(1, objKtbnStrc.strcSelection.strOtherOption, "P8") <> 0 Then
                                If InStr(Trim(objKtbnStrc.strcSelection.strOpSymbol(18)), "Y") = 0 Then
                                    intKtbnStrcSeqNo = 18
                                    strMessageCd = "W0490"
                                    fncCheckSelectOption = False
                                End If
                            End If
                        End If

                        ' J9
                        If InStr(1, objKtbnStrc.strcSelection.strOtherOption, "J9") <> 0 Then
                            If InStr(Trim(objKtbnStrc.strcSelection.strOpSymbol(17)), "J") <> 0 Or _
                               InStr(Trim(objKtbnStrc.strcSelection.strOpSymbol(17)), "K") <> 0 Or _
                               InStr(Trim(objKtbnStrc.strcSelection.strOpSymbol(17)), "L") <> 0 Then
                                intKtbnStrcSeqNo = 17
                                strMessageCd = "W0500"
                                fncCheckSelectOption = False
                            End If
                        End If

                        ' スクレーパ、ロッドパッキンのみフッ素ゴム指定
                        If InStr(1, objKtbnStrc.strcSelection.strOtherOption, "T9") <> 0 Then
                            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(1), "T") <> 0 Then
                                intKtbnStrcSeqNo = 1
                                strMessageCd = "W0420"
                                fncCheckSelectOption = False
                            End If
                        End If
                    End If
                    '2011/01/31 MOD RM1101046(2月VerUP：オプション外特注処理追加) <---END

                    '2010/08/25 ADD RM1008009(9月VerUP：SCPD2シリーズ修正) START --->
                Case "SCPD2", "SCPD2-L", "SCPD2-K", "SCPD2-KL", "SCPD2-M", "SCPD2-ML", "SCPD2-O", "SCPD2-OL", "SCPD2-T", "SCPD2-V", "SCPD2-VL"
                    Dim I1 As Integer = 0
                    Dim I2 As Integer = 0
                    Dim isCheck As Boolean = True

                    Select Case objKtbnStrc.strcSelection.strSeriesKataban
                        Case "SCPD2"
                            I1 = 10
                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                Case "C"
                                    I2 = 1
                                Case ""
                                    I2 = 2
                                Case Else
                                    'チェック対象外
                                    isCheck = False
                            End Select
                        Case "SCPD2-L"
                            I1 = 10
                            Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                Case "C"
                                    I2 = 1
                                Case "4", ""
                                    I2 = 2
                                Case Else
                                    'チェック対象外
                                    isCheck = False
                            End Select
                        Case Else
                            Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 8, 1)
                                Case "L"
                                    I1 = 9
                                Case Else
                                    I1 = 6
                            End Select
                            I2 = 1
                    End Select

                    'チェック実行
                    If isCheck Then
                        If objKtbnStrc.strcSelection.strOpSymbol(I1).IndexOf("B1") >= 0 Then
                            'B1は支持形式CBの場合は無条件で選択可
                            If Trim(objKtbnStrc.strcSelection.strOpSymbol(I2)) = "CB" Then
                            Else
                                If objKtbnStrc.strcSelection.strOpSymbol(I1).IndexOf("Y") < 0 Then
                                    intKtbnStrcSeqNo = I1
                                    strMessageCd = "W2770"
                                    fncCheckSelectOption = False
                                End If
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(I1).IndexOf("B2") >= 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(I1).IndexOf("I") < 0 Then
                                intKtbnStrcSeqNo = I1
                                strMessageCd = "W2780"
                                fncCheckSelectOption = False
                            End If
                        End If
                    End If
                    '2010/08/25 ADD RM1008009(9月VerUP：SCPD2シリーズ修正) <--- END
                Case "JSG", "JSG-V"
                    'JSGのみチェック
                    'RM1305005 2013/05/31 ローカル版との差異修正
                    If Left(objKtbnStrc.strcSelection.strSeriesKataban, 3) = "JSG" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(11).IndexOf("B1") >= 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "CB" Then
                            Else
                                If objKtbnStrc.strcSelection.strOpSymbol(11).IndexOf("Y") < 0 Then
                                    intKtbnStrcSeqNo = 11
                                    strMessageCd = "W0290"
                                    fncCheckSelectOption = False
                                End If
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(11).IndexOf("B2") >= 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "CA" Then
                            Else
                                If objKtbnStrc.strcSelection.strOpSymbol(11).IndexOf("I") < 0 Then
                                    intKtbnStrcSeqNo = 11
                                    strMessageCd = "W0300"
                                    fncCheckSelectOption = False
                                End If
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(11).IndexOf("B3") >= 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "CB" Then
                            Else
                                If objKtbnStrc.strcSelection.strOpSymbol(11).IndexOf("Y") < 0 Then
                                    intKtbnStrcSeqNo = 11
                                    strMessageCd = "W0310"
                                    fncCheckSelectOption = False
                                End If
                            End If
                        End If
                    End If

                    'JSG/JSG-V共通
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim <> "" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            Case "00", "LB", "FA", "FB", "CA", "CB"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                    Case "H", "R"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 10 Then
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "D"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 25 Then
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "T"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 30 Then
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "4"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 35 Then
                                            intKtbnStrcSeqNo = 5
                                            strMessageCd = "W0190"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            Case "TC"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "40"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "H", "R"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 68 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "D"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 68 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "T"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 98 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "4"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 98 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "50"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "H", "R"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 68 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "D"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 68 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "T"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 98 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "4"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 98 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "63"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "H", "R"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 74 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "D"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 74 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "T"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 98 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "4"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 98 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "80"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "H", "R"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 86 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "D"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 86 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "T"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 101 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "4"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 101 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                    Case "100"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "H", "R"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 92 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "D"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 92 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "T"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 107 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                            Case "4"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 107 Then
                                                    intKtbnStrcSeqNo = 5
                                                    strMessageCd = "W0190"
                                                    fncCheckSelectOption = False
                                                End If
                                        End Select
                                End Select
                            Case "TA"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "40"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "H"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                                    Case "T0H", "T0V", "T5H", "T5V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 38 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T8H", "T8V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 41 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T2H", "T2V", "T3H", "T3V", "T2WH", "T2WV", "T3WH", "T3WV"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 32 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T1H", "T1V", "T2YH", "T2YV", "T3YH", "T3YV", "T2YD", "T2YDT", "T2YDU"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 43 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case Else
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 42 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                            Case "T"
                                            Case "4"
                                        End Select
                                    Case "50"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "H"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                                    Case "T0H", "T0V", "T5H", "T5V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 51 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T8H", "T8V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 54 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T2H", "T2V", "T3H", "T3V", "T2WH", "T2WV", "T3WH", "T3WV"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 31 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case Else
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 42 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                            Case "T"
                                            Case "4"
                                        End Select
                                    Case "63"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "H"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                                    Case "T0H", "T0V", "T5H", "T5V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 41 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T8H", "T8V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 44 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T2H", "T2V", "T3H", "T3V", "T2WH", "T2WV", "T3WH", "T3WV"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 37 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case Else
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 48 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                            Case "T"
                                            Case "4"
                                        End Select
                                    Case "80"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "H"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                                    Case "T0H", "T0V", "T5H", "T5V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 41 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T8H", "T8V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 43 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T2H", "T2V", "T3H", "T3V", "T2WH", "T2WV", "T3WH", "T3WV"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 37 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T1H", "T1V", "T2YH", "T2YV", "T3YH", "T3YV", "T2YD", "T2YDT", "T2YDU"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 48 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case Else
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 54 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                            Case "T"
                                            Case "4"
                                        End Select
                                    Case "100"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "H"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                                    Case "T0H", "T0V", "T5H", "T5V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 47 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T8H", "T8V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 49 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T2H", "T2V", "T3H", "T3V", "T2WH", "T2WV", "T3WH", "T3WV"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 43 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T1H", "T1V", "T2YH", "T2YV", "T3YH", "T3YV", "T2YD", "T2YDT", "T2YDU"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 54 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case Else
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 60 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                            Case "T"
                                            Case "4"
                                        End Select
                                End Select
                            Case "TB"
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                    Case "40"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                                    Case "T0H", "T0V", "T5H", "T5V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 38 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T8H", "T8V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 41 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T2H", "T2V", "T3H", "T3V", "T2WH", "T2WV", "T3WH", "T3WV"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 32 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T1H", "T1V", "T2YH", "T2YV", "T3YH", "T3YV", "T2YD", "T2YDT", "T2YDU"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 43 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case Else
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 42 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                            Case "T"
                                            Case "4"
                                        End Select
                                    Case "50"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                                    Case "T0H", "T0V", "T5H", "T5V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 53 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T8H", "T8V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 55 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T2H", "T2V", "T3H", "T3V", "T2WH", "T2WV", "T3WH", "T3WV"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 32 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T1H", "T1V", "T2YH", "T2YV", "T3YH", "T3YV", "T2YD", "T2YDT", "T2YDU"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 43 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case Else
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 42 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                            Case "T"
                                            Case "4"
                                        End Select
                                    Case "63"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                                    Case "T0H", "T0V", "T5H", "T5V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 42 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T8H", "T8V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 44 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T2H", "T2V", "T3H", "T3V", "T2WH", "T2WV", "T3WH", "T3WV"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 38 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T1H", "T1V", "T2YH", "T2YV", "T3YH", "T3YV", "T2YD", "T2YDT", "T2YDU"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 49 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case Else
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 48 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                            Case "T"
                                            Case "4"
                                        End Select
                                    Case "80"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                                    Case "T0H", "T0V", "T5H", "T5V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 47 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T8H", "T8V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 49 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T2H", "T2V", "T3H", "T3V", "T2WH", "T2WV", "T3WH", "T3WV"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 43 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case Else
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 54 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                            Case "T"
                                            Case "4"
                                        End Select
                                    Case "100"
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(9).Trim
                                            Case "R"
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                                    Case "T0H", "T0V", "T5H", "T5V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 53 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T8H", "T8V"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 55 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case "T2H", "T2V", "T3H", "T3V", "T2WH", "T2WV", "T3WH", "T3WV"
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 49 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                    Case Else
                                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) < 60 Then
                                                            intKtbnStrcSeqNo = 5
                                                            strMessageCd = "W0190"
                                                            fncCheckSelectOption = False
                                                        End If
                                                End Select
                                            Case "D"
                                            Case "T"
                                            Case "4"
                                        End Select
                                End Select
                        End Select
                    End If
                Case "MFC", "MFC-L", "MFC-K", "MFC-KL", "MFC-B", _
                     "MFC-BL", "MFC-BK", "MFC-BKL", "MFC-BS", "MFC-BSK"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(4).Trim) Mod 5 <> 0 Then
                        intKtbnStrcSeqNo = 4
                        strMessageCd = "W0510"
                        fncCheckSelectOption = False
                    End If
                    'RM0912039 2009/12/17 Y.Miura チェックしないよう変更
                    'Case "LCM", "LCM-A", "LCM-P", "LCM-R"
                    '    If objKtbnStrc.strcSelection.strOpSymbol(7).IndexOf("F1") >= 0 Or _
                    '       objKtbnStrc.strcSelection.strOpSymbol(7).IndexOf("F2") >= 0 Then
                    '        intKtbnStrcSeqNo = 7
                    '        strMessageCd = "W0520"
                    '        fncCheckSelectOption = False
                    '    End If
                Case "CAC4"
                    If objKtbnStrc.strcSelection.strOpSymbol(14).Trim <> "" Then
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(14), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "Q"
                                    'オプションY/Y1を同時選択していない時は選択不可
                                    If objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("Y") < 0 Then
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0540"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                        Next
                    End If
                Case "UCAC2"
                    If objKtbnStrc.strcSelection.strOpSymbol(13).Trim <> "" Then
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(13), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "Q"
                                    'オプションY/Y1を同時選択していない時は選択不可
                                    If objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("Y") < 0 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0540"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                        Next
                    End If
                Case "CAC3"
                    Dim intFlg As Integer = 0
                    If IsNumeric(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) = False Then
                        '数値以外は不可
                        intFlg = 1
                    End If

                    For intLoopCnt = 1 To Len(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)
                        If Mid(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, intLoopCnt, 1) = CdCst.Sign.Dot Then
                            intFlg = 1
                            Exit For
                        End If
                    Next

                    If intFlg = 1 Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0330"
                        fncCheckSelectOption = False
                    End If

                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < 50 Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0190"
                        fncCheckSelectOption = False
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(13).Trim <> "" Then
                        strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(1), CdCst.Sign.Delimiter.Comma)
                        For intLoopCnt = 0 To strOpArray.Length - 1
                            Select Case strOpArray(intLoopCnt).Trim
                                Case "Q"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(7).Trim
                                        Case "50", "75", "100", "125", "150"
                                        Case Else
                                            intKtbnStrcSeqNo = 7
                                            strMessageCd = "W0530"
                                            fncCheckSelectOption = False
                                    End Select

                                    'オプションY/Y1を同時選択していない時は選択不可
                                    If objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("Y") < 0 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0540"
                                        fncCheckSelectOption = False
                                    End If
                                Case "K"
                                    If objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("Y") < 0 And _
                                       objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("I") < 0 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0550"
                                        fncCheckSelectOption = False
                                    End If
                                Case "D"
                                    If objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("Y") < 0 And _
                                       objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("I") < 0 Then
                                        intKtbnStrcSeqNo = 13
                                        strMessageCd = "W0560"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                        Next
                    End If
                    '↓RM1306005 2013/06/04 追加
                Case "SRM3"
                    '2013/06/19 修正
                    If objKtbnStrc.strcSelection.strKeyKataban = "" Then
                        If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "SX" Then
                            Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(1))
                                Case "12", "16", "20", "25", "32"
                                    If Trim(objKtbnStrc.strcSelection.strOpSymbol(3)) > 1000 Then
                                        intKtbnStrcSeqNo = 3
                                        strMessageCd = "W0200"
                                        fncCheckSelectOption = False
                                    End If
                                Case "40", "50", "63"
                                    If Trim(objKtbnStrc.strcSelection.strOpSymbol(3)) > 500 Then
                                        intKtbnStrcSeqNo = 3
                                        strMessageCd = "W0200"
                                        fncCheckSelectOption = False
                                    End If
                                Case "80"
                                    If Trim(objKtbnStrc.strcSelection.strOpSymbol(3)) > 400 Then
                                        intKtbnStrcSeqNo = 3
                                        strMessageCd = "W0200"
                                        fncCheckSelectOption = False
                                    End If
                                Case "100"
                                    If Trim(objKtbnStrc.strcSelection.strOpSymbol(3)) > 300 Then
                                        intKtbnStrcSeqNo = 3
                                        strMessageCd = "W0200"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                        End If
                    End If
                Case "SRL3"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "", "G", "J"
                            If objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "SX" Then
                                Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(3))
                                    Case "12", "16", "20", "25", "32"
                                        If Trim(objKtbnStrc.strcSelection.strOpSymbol(6)) > 1000 Then
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "40", "50", "63"
                                        If Trim(objKtbnStrc.strcSelection.strOpSymbol(6)) > 500 Then
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "80"
                                        If Trim(objKtbnStrc.strcSelection.strOpSymbol(6)) > 400 Then
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                    Case "100"
                                        If Trim(objKtbnStrc.strcSelection.strOpSymbol(6)) > 300 Then
                                            intKtbnStrcSeqNo = 6
                                            strMessageCd = "W0200"
                                            fncCheckSelectOption = False
                                        End If
                                End Select
                            End If
                    End Select
                    '↑RM1306005 2013/06/04 追加
                Case "PCU2"
                    If InStr(objKtbnStrc.strcSelection.strOpSymbol(7), "B1") <> 0 Then
                        If Trim(objKtbnStrc.strcSelection.strOpSymbol(1)) <> "CB" And _
                           InStr(objKtbnStrc.strcSelection.strOpSymbol(7), "Y") = 0 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W8400"
                        End If
                    End If

                    If InStr(objKtbnStrc.strcSelection.strOpSymbol(7), "B2") <> 0 Then
                        If Trim(objKtbnStrc.strcSelection.strOpSymbol(1)) <> "CA" And _
                           InStr(objKtbnStrc.strcSelection.strOpSymbol(7), "I") = 0 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W8410"
                        End If
                    End If

                Case "STR2-B", "STR2-M"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "D", "Q"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                                Case "6", "10"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) > 50 Then
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W0200"
                                        fncCheckSelectOption = False
                                    End If
                                Case "16", "20", "25", "32"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim) > 100 Then
                                        intKtbnStrcSeqNo = 5
                                        strMessageCd = "W0200"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                    End Select
            End Select

            'RM0906034 2009/08/18 Y.Miura　二次電池対応
            'P4の必須チェック
            Dim bolOptionP4 As Boolean = False
            Dim strOptionP4 As String = String.Empty
            Dim bolOptionU As Boolean = False

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "LCG"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(11), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "U"
                                        bolOptionU = True
                                End Select
                            Next
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(12), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                        strOptionP4 = strOpArray(intLoopCnt).Trim
                                End Select
                            Next
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                Case "20", "25"
                                    If bolOptionU And strOptionP4 = "P40" Then
                                        intKtbnStrcSeqNo = 12
                                        strMessageCd = "W8780"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 12
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "LCG-Q"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "5"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(11), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "U"
                                        bolOptionU = True
                                End Select
                            Next
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(12), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                        strOptionP4 = strOpArray(intLoopCnt).Trim
                                End Select
                            Next
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                                Case "20", "25"
                                    If bolOptionU And strOptionP4 = "P40" Then
                                        intKtbnStrcSeqNo = 12
                                        strMessageCd = "W8780"
                                        fncCheckSelectOption = False
                                    End If
                            End Select
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 12
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "LCS"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                End Select
                            Next
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 10
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "LCR"          'RM1005030 2010/05/25 Y.Miura
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "5"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(12), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                End Select
                            Next
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 12
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "LCR-Q"          'RM1307003 2013/07/04
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(12), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                End Select
                            Next
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 12
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "STG-B", "STG-M"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(9), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                End Select
                            Next
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "SRL3"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4", "R"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                End Select
                            Next
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 10
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "SSD"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(19), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                End Select
                            Next
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 19
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                        Case "E"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(11), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                End Select
                            Next
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 11
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                        Case "P"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(17), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                End Select
                            Next
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 17
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "SSD2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(19), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                End Select
                            Next
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 19
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                        Case "E"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(9), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                End Select
                            Next
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 9
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                        Case "L"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(19), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                End Select
                            Next
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 19
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "USSD", "USSD-L", "USSD-K", "USSD-KL"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "2"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(7), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                End Select
                            Next
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "SCM"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(13), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                End Select
                            Next
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 13
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "SCG", "SCG-Q", "SCG-U"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(10), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4", "P40"
                                        bolOptionP4 = True
                                End Select
                            Next
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 10
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                    End Select

                    'RM0908030 2009/09/08 Y.Miura 二次電池対応機種
                Case "MDC2", "MDC2-L"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            If fncP4Check(objKtbnStrc, _
                                       intKtbnStrcSeqNo, _
                                       strOptionSymbol, _
                                       strMessageCd, _
                                       7) = False Then
                                fncCheckSelectOption = False
                            End If
                    End Select

                    'RM0908030 2009/10/19 Y.Miura 二次電池対応機種
                Case "STK"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            If fncP4Check(objKtbnStrc, _
                                       intKtbnStrcSeqNo, _
                                       strOptionSymbol, _
                                       strMessageCd, _
                                       7) = False Then
                                fncCheckSelectOption = False
                            End If
                    End Select
                    'RM0908030 2009/10/19 Y.Miura 二次電池対応機種
                Case "SRM3", "SRM3-Q"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            If fncP4Check(objKtbnStrc, _
                                       intKtbnStrcSeqNo, _
                                       strOptionSymbol, _
                                       strMessageCd, _
                                       8) = False Then
                                fncCheckSelectOption = False
                            End If
                    End Select
                    'RM0908030 2009/10/20 Y.Miura 二次電池対応機種
                Case "HKP"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            Dim bolOptionG As Boolean = False           'オプション
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(2), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "G"
                                        bolOptionG = True
                                End Select
                            Next
                            If Not bolOptionG Then                      'Gが選択されていないとエラーにする
                                intKtbnStrcSeqNo = 2
                                strMessageCd = "W8800"
                                fncCheckSelectOption = False
                                Exit Select
                            End If
                            'P4,P40が選択されていないとエラーにする
                            If fncP4Check(objKtbnStrc, _
                                       intKtbnStrcSeqNo, _
                                       strOptionSymbol, _
                                       strMessageCd, _
                                       7) = False Then
                                fncCheckSelectOption = False
                            End If
                    End Select
                    'RM0908030 2009/10/20 Y.Miura 二次電池対応機種
                Case "CKG"   'チャック
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            Dim bolOptionG As Boolean = False           'オプション
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(2), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "G"
                                        bolOptionG = True
                                End Select
                            Next
                            If Not bolOptionG Then                      'Gが選択されていないとエラーにする
                                intKtbnStrcSeqNo = 2
                                strMessageCd = "W8800"
                                fncCheckSelectOption = False
                                Exit Select
                            End If
                            'P4,P40が選択されていないとエラーにする
                            If fncP4Check(objKtbnStrc, _
                                       intKtbnStrcSeqNo, _
                                       strOptionSymbol, _
                                       strMessageCd, _
                                       7) = False Then
                                fncCheckSelectOption = False
                            End If
                    End Select
                    'RM1001045 2010/02/24 Y.Miura 二次電池対応機種
                Case "SCPD2-L"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            If Not fncP4Check(objKtbnStrc, _
                                       intKtbnStrcSeqNo, _
                                       strOptionSymbol, _
                                       strMessageCd, _
                                       9) Then
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "SCPG2-L"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            If Not fncP4Check(objKtbnStrc, _
                                       intKtbnStrcSeqNo, _
                                       strOptionSymbol, _
                                       strMessageCd, _
                                       9) Then
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "MRL2-L", "MRL2-GL", "MRL2-WL"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            If Not fncP4Check(objKtbnStrc, _
                                       intKtbnStrcSeqNo, _
                                       strOptionSymbol, _
                                       strMessageCd, _
                                       8) Then
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "MRG2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            If Not fncP4Check(objKtbnStrc, _
                                       intKtbnStrcSeqNo, _
                                       strOptionSymbol, _
                                       strMessageCd, _
                                       6) Then
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "GRC", "GRC-K"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            If Not fncP4Check(objKtbnStrc, _
                                       intKtbnStrcSeqNo, _
                                       strOptionSymbol, _
                                       strMessageCd, _
                                       7) Then
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "UCA2", "UCA2-B"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            If Not fncP4Check(objKtbnStrc, _
                                       intKtbnStrcSeqNo, _
                                       strOptionSymbol, _
                                       strMessageCd, _
                                       4) Then
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "UCA2-L", "UCA2-BL"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            If Not fncP4Check(objKtbnStrc, _
                                       intKtbnStrcSeqNo, _
                                       strOptionSymbol, _
                                       strMessageCd, _
                                       7) Then
                                fncCheckSelectOption = False
                            End If
                    End Select
                Case "SCS2"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "4"
                            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(17), CdCst.Sign.Delimiter.Comma)
                            For intLoopCnt = 0 To strOpArray.Length - 1
                                Select Case strOpArray(intLoopCnt).Trim
                                    Case "P4"
                                        bolOptionP4 = True
                                End Select
                            Next
                            If Not bolOptionP4 Then
                                intKtbnStrcSeqNo = 17
                                strMessageCd = "W8770"
                                fncCheckSelectOption = False
                            End If
                    End Select

                Case Else



            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncP4Check
    '*【処理】
    '*  二次電池対応機器チェック
    '*【概要】
    '*  二次電池が含まれるかをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*  <Integer>      intOptionPos         要素位置　　　　　   
    '*【戻り値】
    '*  <Boolean>
    '*【更新】
    '*  ・受付No：RM0906034  二次電池対応機器対応　新規追加
    '*                                      更新日：2009/09/08   更新者：Y.Miura
    '********************************************************************************************
    Private Function fncP4Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String, _
                                          ByVal intOptionPos As Integer) As Boolean

        Try

            fncP4Check = True

            '二次電池対応
            Dim bolOpP4 As Boolean = False
            Dim strOpArray() As String
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt As Integer = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "P4", "P40"
                        bolOpP4 = True
                End Select
            Next
            'P4の必須チェック
            If Not bolOpP4 Then
                intKtbnStrcSeqNo = intOptionPos
                strMessageCd = "W8770"
                fncP4Check = False
                Exit Try
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncGetMinStroke_CAV2
    '*【処理】
    '*  CAV2最小ストローク取得
    '*【概要】
    '*  CAV2の最小ストロークを取得する
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*【戻り値】
    '*  <integer>
    '*【更新】
    '*  ・受付No：RM0911XXX  最小ストローク取得を関数化
    '*                                      更新日：2009/11/10   更新者：Y.Miura
    '********************************************************************************************
    Private Function fncGetMinStroke_CAV2(ByVal objKtbnStrc As KHKtbnStrc) As Integer

        Dim strKisyu As String          '機種
        Dim intStroke As Integer        'ストローク
        Dim strSiji As String           '支持形式
        Dim strBore As String           '内径
        Dim strCushion As String        'クッション
        Dim strSwitch As String         'スイッチ形番
        Dim strSwitchNum As String      'スイッチ数
        Dim intSwitchNum As Integer
        Dim intMin As Integer = 0

        Try
            strKisyu = objKtbnStrc.strcSelection.strSeriesKataban.Trim          '機種
            strSiji = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '支持形式
            strBore = objKtbnStrc.strcSelection.strOpSymbol(3).Trim             '内径
            strCushion = objKtbnStrc.strcSelection.strOpSymbol(4).Trim          'クッション
            intStroke = CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim)     'ストローク
            strSwitch = objKtbnStrc.strcSelection.strOpSymbol(7).Trim           'SW形番
            strSwitchNum = objKtbnStrc.strcSelection.strOpSymbol(9).Trim        'SW個数記号

            'スイッチ数
            Select Case strSwitchNum
                Case "R", "H" : intSwitchNum = 1
                Case "D" : intSwitchNum = 2
                Case "T" : intSwitchNum = 3
                Case "4" : intSwitchNum = 4
                Case Else
                    intSwitchNum = 1
            End Select

            Select Case strKisyu    '機種
                Case "CAV2", "COVN2", "COVP2"
                    Select Case strSiji '支持形式
                        Case "TC", "TF"
                            Select Case Trim(strSwitch)     'SW形番
                                Case "T0H", "T5H", "T8H"    'SW ｽﾄﾚｰﾄﾀｲﾌﾟ
                                    Select Case strCushion   'クッション
                                        Case "N"                'クッションなし
                                            Select Case strBore
                                                Case "50" : intMin = 215
                                                Case "75" : intMin = 193
                                                Case "100" : intMin = 83
                                            End Select
                                        Case "B"                'クッション付
                                            Select Case strBore
                                                Case "50" : intMin = 241
                                                Case "75" : intMin = 241
                                                Case "100" : intMin = 120
                                            End Select
                                    End Select
                                Case "T0V", "T5V", "T8V"    'SW Ｌ字タイプ
                                    Select Case strCushion   'クッション
                                        Case "N"                'クッションなし
                                            Select Case strBore
                                                Case "50" : intMin = 215
                                                Case "75" : intMin = 193
                                                Case "100"
                                                    Select Case intSwitchNum
                                                        Case 1 : intMin = 71
                                                        Case 2 : intMin = 71
                                                        Case 3 : intMin = 73
                                                    End Select
                                            End Select
                                        Case "B"                'クッション付
                                            Select Case strBore
                                                Case "50" : intMin = 241
                                                Case "75" : intMin = 241
                                                Case "100"
                                                    Select Case intSwitchNum
                                                        Case 1 : intMin = 108
                                                        Case 2 : intMin = 108
                                                        Case 3 : intMin = 110
                                                    End Select
                                            End Select
                                    End Select
                                Case "T1H", "T2H", "T3H", "T2YH", "T3YH", "T2JH", "T2WH", "T3WH"
                                    Select Case strCushion   'クッション
                                        Case "N"                'クッションなし
                                            Select Case strBore
                                                Case "50" : intMin = 76
                                                Case "75" : intMin = 54
                                                Case "100" : intMin = 84
                                            End Select
                                        Case "B"                'クッション付
                                            Select Case strBore
                                                Case "50" : intMin = 102
                                                Case "75" : intMin = 102
                                                Case "100" : intMin = 121
                                            End Select
                                    End Select
                                Case "T1V", "T2V", "T3V", "T2YV", "T3YV", "T2JV", "T2WV", "T3WV"
                                    Select Case strCushion   'クッション
                                        Case "N"                'クッションなし
                                            Select Case strBore
                                                Case "50"
                                                    Select Case intSwitchNum
                                                        Case 1 : intMin = 46
                                                        Case 2 : intMin = 46
                                                        Case 3 : intMin = 47
                                                    End Select
                                                Case "75"
                                                    Select Case intSwitchNum
                                                        Case 1 : intMin = 24
                                                        Case 2 : intMin = 24
                                                        Case 3 : intMin = 26
                                                    End Select
                                                Case "100"
                                                    Select Case intSwitchNum
                                                        Case 1 : intMin = 54
                                                        Case 2 : intMin = 54
                                                        Case 3 : intMin = 58
                                                    End Select
                                            End Select
                                        Case "B"                'クッション付
                                            Select Case strBore
                                                Case "50"
                                                    Select Case intSwitchNum
                                                        Case 1 : intMin = 72
                                                        Case 2 : intMin = 72
                                                        Case 3 : intMin = 73
                                                    End Select
                                                Case "75"
                                                    Select Case intSwitchNum
                                                        Case 1 : intMin = 72
                                                        Case 2 : intMin = 72
                                                        Case 3 : intMin = 74
                                                    End Select
                                                Case "100"
                                                    Select Case intSwitchNum
                                                        Case 1 : intMin = 91
                                                        Case 2 : intMin = 91
                                                        Case 3 : intMin = 95
                                                    End Select
                                            End Select
                                    End Select
                                Case Else

                            End Select
                        Case Else

                    End Select
            End Select

        Catch ex As Exception

            Throw ex
        Finally
            fncGetMinStroke_CAV2 = intMin
        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncGetMinStroke_SCG
    '*【処理】
    '*  SCG最小ストローク取得
    '*【概要】
    '*  CAV2の最小ストロークを取得する
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*【戻り値】
    '*  <integer>
    '*【更新】
    '*  ・受付No：RM0911XXX  最小ストローク取得を関数化
    '*                                      更新日：2009/11/10   更新者：Y.Miura
    '********************************************************************************************
    Private Function fncGetMinStroke_SCG(ByVal objKtbnStrc As KHKtbnStrc) As Integer

        Dim strKisyu As String          '機種
        Dim intStroke As Integer        'ストローク
        Dim strSiji As String           '支持形式
        Dim strBore As String           '内径
        Dim strSwitch As String         'スイッチ形番
        Dim strSwitchNum As String      'スイッチ数
        Dim intSwitchNum As Integer
        Dim strSwitchType As String
        Dim intMin As Integer = 0

        Try
            strKisyu = objKtbnStrc.strcSelection.strSeriesKataban.Trim          '機種
            strSiji = objKtbnStrc.strcSelection.strOpSymbol(1).Trim             '支持形式
            strBore = objKtbnStrc.strcSelection.strOpSymbol(2).Trim             '内径
            intStroke = CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim)     'ストローク
            strSwitch = fncConvertSwitchP4ToStandard(objKtbnStrc.strcSelection.strOpSymbol(7).Trim)     'SW形番
            strSwitchType = Right(strSwitch, 1)                                 'V:リード線Ｌ字タイプ
            strSwitchNum = objKtbnStrc.strcSelection.strOpSymbol(9).Trim        'SW個数記号
            Select Case strSwitchNum
                Case "R", "H" : intSwitchNum = 1
                Case "D" : intSwitchNum = 2
                Case "T" : intSwitchNum = 3
                Case "4" : intSwitchNum = 4
                Case Else
                    intSwitchNum = 1
            End Select

            Select Case Left(strSwitch, 3)
                Case "T0H", "T0V", "T5H", "T5V"
                    Select Case strSiji
                        Case "00", "LB", "FA", "FB", "CA", "CB"
                            Select Case strBore
                                Case "32"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 9
                                        Case 2 : intMin = 17
                                        Case 3 : intMin = 34
                                        Case 4 : intMin = 51
                                    End Select
                                Case "40"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 9
                                        Case 2 : intMin = 18
                                        Case 3 : intMin = 36
                                        Case 4 : intMin = 54
                                    End Select
                                Case "50"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 9
                                        Case 2 : intMin = 18
                                        Case 3 : intMin = 36
                                        Case 4 : intMin = 54
                                    End Select
                                Case "63"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 10
                                        Case 2 : intMin = 19
                                        Case 3 : intMin = 38
                                        Case 4 : intMin = 57
                                    End Select
                                Case "80"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 10
                                        Case 2 : intMin = 20
                                        Case 3 : intMin = 39
                                        Case 4 : intMin = 59
                                    End Select
                                Case "100"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 10
                                        Case 2 : intMin = 20
                                        Case 3 : intMin = 40
                                        Case 4 : intMin = 60
                                    End Select
                            End Select
                        Case "TC"       'ＴＣ　中間トラニオン形
                            Select Case strBore
                                Case "32"   'φ３２
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 94
                                        Case 2 : intMin = 94
                                        Case 3
                                            Select Case strSwitchType
                                                Case "H" : intMin = 169
                                                Case "V" : intMin = 155
                                            End Select
                                        Case 4
                                            Select Case strSwitchType
                                                Case "H" : intMin = 169
                                                Case "V" : intMin = 155
                                            End Select
                                    End Select
                                Case "40"   'φ４０
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 81
                                        Case 2 : intMin = 81
                                        Case 3
                                            Select Case strSwitchType
                                                Case "H" : intMin = 164
                                                Case "V" : intMin = 142
                                            End Select
                                        Case 4
                                            Select Case strSwitchType
                                                Case "H" : intMin = 164
                                                Case "V" : intMin = 142
                                            End Select
                                    End Select
                                Case "50"   'φ５０
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 112
                                        Case 2 : intMin = 112
                                        Case 3 : intMin = 121
                                        Case 4 : intMin = 121
                                    End Select
                                Case "63"   'φ６３
                                    Select Case intSwitchNum
                                        Case 1
                                            Select Case strSwitchType
                                                Case "H" : intMin = 85
                                                Case "V" : intMin = 73
                                            End Select
                                        Case 2
                                            Select Case strSwitchType
                                                Case "H" : intMin = 85
                                                Case "V" : intMin = 73
                                            End Select
                                        Case 3 : intMin = 91
                                        Case 4 : intMin = 91
                                    End Select
                                Case "80"   'φ８０
                                    Select Case intSwitchNum
                                        Case 1
                                            Select Case strSwitchType
                                                Case "H" : intMin = 96
                                                Case "V" : intMin = 79
                                            End Select
                                        Case 2
                                            Select Case strSwitchType
                                                Case "H" : intMin = 96
                                                Case "V" : intMin = 79
                                            End Select
                                        Case 3 : intMin = 99
                                        Case 4 : intMin = 99
                                    End Select
                                Case "100"  'φ１００
                                    Select Case intSwitchNum
                                        Case 1
                                            Select Case strSwitchType
                                                Case "H" : intMin = 101
                                                Case "V" : intMin = 84
                                            End Select
                                        Case 2
                                            Select Case strSwitchType
                                                Case "H" : intMin = 101
                                                Case "V" : intMin = 84
                                            End Select
                                        Case 3 : intMin = 105
                                        Case 4 : intMin = 105
                                    End Select
                            End Select
                        Case "TA"   'ＴＡ　ロッド側トラニオン形
                            Select Case strBore
                                Case 32 : intMin = 42
                                Case 40 : intMin = 38
                                Case 50 : intMin = 51
                                Case 63 : intMin = 41
                                Case 80 : intMin = 41
                                Case 100 : intMin = 47
                            End Select
                        Case "TB"   'ＴＢ　ヘッド側トラニオン形
                            Select Case strBore
                                Case 32 : intMin = 42
                                Case 40 : intMin = 38
                                Case 50 : intMin = 53
                                Case 63 : intMin = 42
                                Case 80 : intMin = 47
                                Case 100 : intMin = 53
                            End Select
                    End Select
                Case "T8H", "T8V"
                    Select Case strSiji
                        Case "00", "LB", "FA", "FB", "CA", "CB"
                            Select Case strBore
                                Case "32"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 9
                                        Case 2 : intMin = 17
                                        Case 3 : intMin = 34
                                        Case 4 : intMin = 51
                                    End Select
                                Case "40"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 9
                                        Case 2 : intMin = 18
                                        Case 3 : intMin = 36
                                        Case 4 : intMin = 54
                                    End Select
                                Case "50"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 9
                                        Case 2 : intMin = 18
                                        Case 3 : intMin = 36
                                        Case 4 : intMin = 54
                                    End Select
                                Case "63"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 10
                                        Case 2 : intMin = 19
                                        Case 3 : intMin = 38
                                        Case 4 : intMin = 57
                                    End Select
                                Case "80"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 10
                                        Case 2 : intMin = 20
                                        Case 3 : intMin = 39
                                        Case 4 : intMin = 59
                                    End Select
                                Case "100"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 10
                                        Case 2 : intMin = 20
                                        Case 3 : intMin = 40
                                        Case 4 : intMin = 60
                                    End Select
                            End Select
                        Case "TC"
                            Select Case strBore
                                Case "32"
                                    Select Case intSwitchNum    'SW数
                                        Case 1 : intMin = 100
                                        Case 2 : intMin = 100
                                        Case 3
                                            Select Case strSwitchType
                                                Case "H" : intMin = 191
                                                Case "V" : intMin = 161
                                            End Select
                                        Case 4
                                            Select Case strSwitchType
                                                Case "H" : intMin = 191
                                                Case "V" : intMin = 161
                                            End Select
                                    End Select
                                Case "40"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 87
                                        Case 2 : intMin = 87
                                        Case 3
                                            Select Case strSwitchType
                                                Case "H" : intMin = 178
                                                Case "V" : intMin = 148
                                            End Select
                                        Case 4
                                            Select Case strSwitchType
                                                Case "H" : intMin = 178
                                                Case "V" : intMin = 148
                                            End Select
                                    End Select
                                Case "50"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 116
                                        Case 2 : intMin = 116
                                        Case 3 : intMin = 121
                                        Case 4 : intMin = 121
                                    End Select
                                Case "63"
                                    Select Case intSwitchNum
                                        Case 1
                                            Select Case strSwitchType
                                                Case "H" : intMin = 89
                                                Case "V" : intMin = 77
                                            End Select
                                        Case 2
                                            Select Case strSwitchType
                                                Case "H" : intMin = 89
                                                Case "V" : intMin = 77
                                            End Select
                                        Case 3 : intMin = 99
                                        Case 4 : intMin = 99
                                    End Select
                                Case "80"
                                    Select Case intSwitchNum
                                        Case 1
                                            Select Case strSwitchType
                                                Case "H" : intMin = 100
                                                Case "V" : intMin = 75
                                            End Select
                                        Case 2
                                            Select Case strSwitchType
                                                Case "H" : intMin = 100
                                                Case "V" : intMin = 75
                                            End Select
                                        Case 3 : intMin = 111
                                        Case 4 : intMin = 111
                                    End Select
                                Case "100"
                                    Select Case intSwitchNum
                                        Case 1
                                            Select Case strSwitchType
                                                Case "H" : intMin = 105
                                                Case "V" : intMin = 80
                                            End Select
                                        Case 2
                                            Select Case strSwitchType
                                                Case "H" : intMin = 105
                                                Case "V" : intMin = 80
                                            End Select
                                        Case 3 : intMin = 117
                                        Case 4 : intMin = 117
                                    End Select
                            End Select
                        Case "TA"
                            Select Case strBore
                                Case 32 : intMin = 45
                                Case 40 : intMin = 41
                                Case 50 : intMin = 54
                                Case 63 : intMin = 44
                                Case 80 : intMin = 43
                                Case 100 : intMin = 49
                            End Select
                        Case "TB"
                            Select Case strBore
                                Case 32 : intMin = 45
                                Case 40 : intMin = 41
                                Case 50 : intMin = 55
                                Case 63 : intMin = 44
                                Case 80 : intMin = 49
                                Case 100 : intMin = 55
                            End Select
                    End Select
                Case "T1H", "T1V", "T2Y", "T3Y", "T2W", "T3W", "T2J"
                    Select Case strSiji
                        Case "00", "LB", "FA", "FB", "CA", "CB"
                            Select Case strBore
                                Case "32"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 6
                                        Case 2 : intMin = 11
                                        Case 3 : intMin = 22
                                        Case 4 : intMin = 33
                                    End Select
                                Case "40"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 6
                                        Case 2 : intMin = 11
                                        Case 3 : intMin = 22
                                        Case 4 : intMin = 33
                                    End Select
                                Case "50"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 6
                                        Case 2 : intMin = 12
                                        Case 3 : intMin = 24
                                        Case 4 : intMin = 36
                                    End Select
                                Case "63"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 6
                                        Case 2 : intMin = 12
                                        Case 3 : intMin = 24
                                        Case 4 : intMin = 36
                                    End Select
                                Case "80"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 7
                                        Case 2 : intMin = 13
                                        Case 3 : intMin = 25
                                        Case 4 : intMin = 38
                                    End Select
                                Case "100"
                                    Select Case intSwitchNum
                                        Case 1 : intMin = 7
                                        Case 2 : intMin = 13
                                        Case 3 : intMin = 26
                                        Case 4 : intMin = 39
                                    End Select
                            End Select
                        Case "TC"   'ＴＣ　中間トラニオン形
                            Select Case strBore
                                Case "32"   'φ３２
                                    Select Case intSwitchNum    'SW数
                                        Case 1
                                            Select Case strSwitchType
                                                Case "H", "D", "T", "U" : intMin = 86
                                                Case "V" : intMin = 61
                                            End Select
                                        Case 2
                                            Select Case strSwitchType
                                                Case "H", "D", "T", "U" : intMin = 86
                                                Case "V" : intMin = 61
                                            End Select
                                        Case 3
                                            Select Case strSwitchType
                                                Case "H", "D", "T", "U" : intMin = 177
                                                Case "V" : intMin = 122
                                            End Select
                                        Case 4
                                            Select Case strSwitchType
                                                Case "H", "D", "T", "U" : intMin = 177
                                                Case "V" : intMin = 122
                                            End Select
                                    End Select
                                Case "40"   'φ４０
                                    Select Case intSwitchNum    'SW数
                                        Case 1
                                            Select Case strSwitchType
                                                Case "H", "D", "T", "U" : intMin = 91
                                                Case "V" : intMin = 66
                                            End Select
                                        Case 2
                                            Select Case strSwitchType
                                                Case "H", "D", "T", "U" : intMin = 91
                                                Case "V" : intMin = 66
                                            End Select
                                        Case 3
                                            Select Case strSwitchType
                                                Case "H", "D", "T", "U" : intMin = 182
                                                Case "V" : intMin = 127
                                            End Select
                                        Case 4
                                            Select Case strSwitchType
                                                Case "H", "D", "T", "U" : intMin = 182
                                                Case "V" : intMin = 127
                                            End Select
                                    End Select
                                Case "50"   'φ５０
                                    Select Case strSwitchType
                                        Case "H", "D", "T", "U" : intMin = 93
                                        Case "V" : intMin = 68
                                    End Select
                                Case "63"   'φ６３
                                    Select Case strSwitchType
                                        Case "H", "D", "T", "U" : intMin = 99
                                        Case "V" : intMin = 74
                                    End Select
                                Case "80"   'φ８０
                                    Select Case intSwitchNum
                                        Case 1 To 2
                                            Select Case strSwitchType
                                                Case "H", "D", "T", "U" : intMin = 110
                                                Case "V" : intMin = 85
                                            End Select
                                        Case 3 To 4
                                            Select Case strSwitchType
                                                Case "H", "D", "T", "U" : intMin = 110
                                                Case "V" : intMin = 86
                                            End Select
                                    End Select
                                Case "100"  'φ１００
                                    Select Case intSwitchNum
                                        Case 1 To 2
                                            Select Case strSwitchType
                                                Case "H", "D", "T", "U" : intMin = 115
                                                Case "V" : intMin = 90
                                            End Select
                                        Case 3 To 4
                                            Select Case strSwitchType
                                                Case "H", "D", "T", "U" : intMin = 115
                                                Case "V" : intMin = 92
                                            End Select
                                    End Select
                            End Select
                        Case "TA"
                            Select Case strBore
                                Case 32 : intMin = 38
                                Case 40 : intMin = 43
                                Case 50 : intMin = 42
                                Case 63 : intMin = 48
                                Case 80 : intMin = 48
                                Case 100 : intMin = 54
                            End Select
                        Case "TB"
                            Select Case strBore
                                Case 32 : intMin = 38
                                Case 40 : intMin = 43
                                Case 50 : intMin = 43
                                Case 63 : intMin = 49
                                Case 80 : intMin = 54
                                Case 100 : intMin = 60
                            End Select
                    End Select

                Case Else
                    Select Case Left(strSwitch, 2)
                        Case "T2", "T3"
                            Select Case strSiji
                                Case "00", "LB", "FA", "FB", "CA", "CB"
                                    Select Case strBore
                                        Case "32", "40", "50"
                                            Select Case intSwitchNum
                                                Case 1 : intMin = 5
                                                Case 2 : intMin = 10
                                                Case 3 : intMin = 20
                                                Case 4 : intMin = 30
                                            End Select
                                        Case "63"
                                            Select Case intSwitchNum
                                                Case 1 : intMin = 6
                                                Case 2 : intMin = 11
                                                Case 3 : intMin = 21
                                                Case 4 : intMin = 32
                                            End Select
                                        Case "80", "100"
                                            Select Case intSwitchNum
                                                Case 1 : intMin = 6
                                                Case 2 : intMin = 11
                                                Case 3 : intMin = 22
                                                Case 4 : intMin = 33
                                            End Select
                                    End Select
                                Case "TC"   'ＴＣ　中間トラニオン形
                                    Select Case strBore
                                        Case "32"
                                            Select Case intSwitchNum
                                                Case 1
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 64
                                                        Case "V" : intMin = 55
                                                    End Select
                                                Case 2
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 64
                                                        Case "V" : intMin = 55
                                                    End Select
                                                Case 3
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 131
                                                        Case "V" : intMin = 116
                                                    End Select
                                                Case 4
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 131
                                                        Case "V" : intMin = 116
                                                    End Select
                                            End Select
                                        Case "40"
                                            Select Case intSwitchNum
                                                Case 1
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 69
                                                        Case "V" : intMin = 60
                                                    End Select
                                                Case 2
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 69
                                                        Case "V" : intMin = 60
                                                    End Select
                                                Case 3
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 152
                                                        Case "V" : intMin = 121
                                                    End Select
                                                Case 4
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 152
                                                        Case "V" : intMin = 121
                                                    End Select
                                            End Select
                                        Case "50"
                                            Select Case intSwitchNum
                                                Case 1
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 71
                                                        Case "V" : intMin = 62
                                                    End Select
                                                Case 2
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 71
                                                        Case "V" : intMin = 62
                                                    End Select
                                                Case 3
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 71
                                                        Case "V" : intMin = 61
                                                    End Select
                                                Case 4
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 71
                                                        Case "V" : intMin = 61
                                                    End Select
                                            End Select
                                        Case "63"
                                            Select Case intSwitchNum
                                                Case 1 To 4
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 77
                                                        Case "V" : intMin = 68
                                                    End Select
                                            End Select
                                        Case "80"
                                            Select Case intSwitchNum
                                                Case 1 To 2
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 88
                                                        Case "V" : intMin = 79
                                                    End Select
                                                Case 3 To 4
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 88
                                                        Case "V" : intMin = 80
                                                    End Select
                                            End Select
                                        Case "100"
                                            Select Case intSwitchNum
                                                Case 1 To 2
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 93
                                                        Case "V" : intMin = 84
                                                    End Select
                                                Case 3 To 4
                                                    Select Case strSwitchType
                                                        Case "H" : intMin = 93
                                                        Case "V" : intMin = 85
                                                    End Select
                                            End Select
                                    End Select
                                Case "TA"
                                    Select Case strBore
                                        Case 32 : intMin = 27
                                        Case 40 : intMin = 32
                                        Case 50 : intMin = 31
                                        Case 63 : intMin = 37
                                        Case 80 : intMin = 37
                                        Case 100 : intMin = 43
                                    End Select
                                Case "TB"
                                    Select Case strBore
                                        Case 32 : intMin = 27
                                        Case 40 : intMin = 32
                                        Case 50 : intMin = 32
                                        Case 63 : intMin = 38
                                        Case 80 : intMin = 43
                                        Case 100 : intMin = 49
                                    End Select
                            End Select
                        Case Else
                    End Select
            End Select

        Catch ex As Exception

            Throw ex
        Finally
            fncGetMinStroke_SCG = intMin
        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncConvertSwitchP4ToStandard
    '*【処理】
    '*  二次電池用スイッチ形番を標準形番に変換する
    '*【概要】
    '*  CAV2の最小ストロークを取得する
    '*【引数】
    '*  <Object>       strSwitchP4          二次電池スイッチ形番
    '*【戻り値】
    '*  <String>
    '*【更新】
    '*  ・受付No：RM0911XXX      更新日：2009/11/10   更新者：Y.Miura
    '********************************************************************************************
    Private Function fncConvertSwitchP4ToStandard(ByVal strSwitchP4 As String) As String

        Dim strSwitch As String = String.Empty

        Try
            Select Case strSwitchP4
                Case "SW11", "SW12", "SW13", "SW18" : strSwitch = "T2H"
                Case "SW14", "SW15", "SW16" : strSwitch = "T2V"
                Case "SW21", "SW22", "SW23" : strSwitch = "T3H"
                Case "SW24", "SW25", "SW26" : strSwitch = "T3V"
                Case "SW27" : strSwitch = "T5H"
                Case "SW31", "SW32", "SW33" : strSwitch = "T2YH"
                Case "SW34", "SW35", "SW36", "SW38" : strSwitch = "T2YV"
                Case "SW37", "SW48" : strSwitch = "T2WV"
                Case "SW40", "SW47", "SW39" : strSwitch = "T2WH"
                Case "SW41", "SW42", "SW43" : strSwitch = "T3YH"
                Case "SW44", "SW45", "SW46" : strSwitch = "T3YV"
                Case "SW51", "SW52", "SW53" : strSwitch = "K2H"
                Case "SW54", "SW55", "SW56" : strSwitch = "K2V"
                Case "SW58" : strSwitch = "K5H"
                Case "SW61", "SW62", "SW63" : strSwitch = "K3H"
                Case "SW64", "SW65", "SW66" : strSwitch = "K3V"
                Case "SW71", "SW72", "SW73" : strSwitch = "M2V"
                Case "SW74", "SW75", "SW76" : strSwitch = "M3V"
                Case "SW81", "SW82", "SW89" : strSwitch = "F2H"
                Case "SW83", "SW84" : strSwitch = "F2V"
                Case "SW85", "SW86" : strSwitch = "F3H"
                Case "SW87", "SW88" : strSwitch = "F3V"
                Case "SW94", "SW95", "SW96" : strSwitch = "M2H"
                Case "SW97", "SW98", "SW99" : strSwitch = "M3H"
                Case Else
                    strSwitch = strSwitchP4
            End Select
        Catch ex As Exception
            Throw ex
        Finally
            fncConvertSwitchP4ToStandard = strSwitch
        End Try

    End Function
End Module
