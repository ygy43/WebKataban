Module KHCylinderJSC3Check

    '********************************************************************************************
    '*【関数名】
    '*  fncCheckSelectOption
    '*【処理】
    '*  シリンダチェック
    '*【概要】
    '*  シリンダＪＳＣ３シリーズをチェックする
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

        Try

            fncCheckSelectOption = True

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "JSC3"
                    '基本ベース毎にチェック
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "1", "R", "S"
                            'φ40～100ベース
                            If fncSmallBoreBaseCheck(objKtbnStrc, _
                                                     intKtbnStrcSeqNo, _
                                                     strOptionSymbol, _
                                                     strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case "2"
                            'φ125～180ベース
                            If fncBigBoreBaseCheck(objKtbnStrc, _
                                                   intKtbnStrcSeqNo, _
                                                   strOptionSymbol, _
                                                   strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                    End Select
                    '↓RM1302XXX 2013/02/04 Y.Tachi
                Case "JSC4"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "2"
                            'φ125～180ベース
                            If fncBigBoreBaseCheck(objKtbnStrc, _
                                                   intKtbnStrcSeqNo, _
                                                   strOptionSymbol, _
                                                   strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case Else
                            If fncStandardBoreBaseCheck(objKtbnStrc, _
                                           intKtbnStrcSeqNo, _
                                           strOptionSymbol, _
                                           strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                    End Select
                  
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSmallBoreBaseCheck
    '*【処理】
    '*  φ40～100ベースチェック
    '*【概要】
    '*  φ40～100ベースをチェックする
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
    '********************************************************************************************
    Public Function fncSmallBoreBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncSmallBoreBaseCheck = True

            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                '最小値チェック
                'バリエーションによる分類
                Case "", "V", "H", "T2", "G", "G1", "VG", "VG1"
                    If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "" Then
                        'スイッチが選択されていない時
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 1 Then
                            intKtbnStrcSeqNo = 8
                            strMessageCd = "W0200"
                            fncSmallBoreBaseCheck = False
                            Exit Try
                        End If
                    Else
                        If objKtbnStrc.strcSelection.strOpSymbol(10).Trim <> "E0" Then
                            'スイッチが"E0"でない時
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(10).Trim
                                Case "R1", "R2", "R2Y", "R3", "R3Y", _
                                     "R0", "R4", "R5", "R6", "H0", "H0Y"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                        Case "00", "LB", "FA", "FB", "FC", _
                                             "CA", "CB"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 10 Then
                                                        intKtbnStrcSeqNo = 8
                                                        strMessageCd = "W0190"
                                                        fncSmallBoreBaseCheck = False
                                                        Exit Try
                                                    End If
                                                Case "D"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 20 Then
                                                        intKtbnStrcSeqNo = 8
                                                        strMessageCd = "W0190"
                                                        fncSmallBoreBaseCheck = False
                                                        Exit Try
                                                    End If
                                                Case "T"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 35 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50", "63", "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 50 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50", "63", "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 55 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TC", "TF"
                                            'トラニオン位置の指定有無
                                            If InStr(1, strOptionSymbol, "AQ") <> 0 Then
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 1 Then
                                                    intKtbnStrcSeqNo = 8
                                                    strMessageCd = "W0190"
                                                    fncSmallBoreBaseCheck = False
                                                    Exit Try
                                                End If
                                            Else
                                                If objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "B" Then
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                        Case "H", "R", "D"
                                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                                Case "40", "50"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 66 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                                Case "63"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 71 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                                Case "80"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 76 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                                Case "100"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 86 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                            End Select
                                                        Case "T", "4"
                                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                                Case "40", "50"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 92 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                                Case "63"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 97 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                                Case "80"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 102 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                                Case "100"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 112 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                            End Select
                                                    End Select
                                                Else
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                        Case "H", "R", "D"
                                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                                Case "40", "50"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 86 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                                Case "63"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 91 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                                Case "80"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 96 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                                Case "100"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 106 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                            End Select
                                                        Case "T", "4"
                                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                                Case "40", "50"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 92 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                                Case "63"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 97 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                                Case "80"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 102 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                                Case "100"
                                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 112 Then
                                                                        intKtbnStrcSeqNo = 8
                                                                        strMessageCd = "W0190"
                                                                        fncSmallBoreBaseCheck = False
                                                                        Exit Try
                                                                    End If
                                                            End Select
                                                    End Select
                                                End If
                                            End If
                                        Case "TA", "TD", "TB", "TE"
                                            If objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "B" Then
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                    Case "H", "R"
                                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                            Case "40"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 28 Then
                                                                    intKtbnStrcSeqNo = 8
                                                                    strMessageCd = "W0190"
                                                                    fncSmallBoreBaseCheck = False
                                                                    Exit Try
                                                                End If
                                                            Case "50"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 26 Then
                                                                    intKtbnStrcSeqNo = 8
                                                                    strMessageCd = "W0190"
                                                                    fncSmallBoreBaseCheck = False
                                                                    Exit Try
                                                                End If
                                                            Case "63"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 31 Then
                                                                    intKtbnStrcSeqNo = 8
                                                                    strMessageCd = "W0190"
                                                                    fncSmallBoreBaseCheck = False
                                                                    Exit Try
                                                                End If
                                                            Case "80"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 34 Then
                                                                    intKtbnStrcSeqNo = 8
                                                                    strMessageCd = "W0190"
                                                                    fncSmallBoreBaseCheck = False
                                                                    Exit Try
                                                                End If
                                                            Case "100"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                                                    intKtbnStrcSeqNo = 8
                                                                    strMessageCd = "W0190"
                                                                    fncSmallBoreBaseCheck = False
                                                                    Exit Try
                                                                End If
                                                        End Select
                                                End Select
                                            Else
                                                Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                    Case "H", "R"
                                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                            Case "40"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 38 Then
                                                                    intKtbnStrcSeqNo = 8
                                                                    strMessageCd = "W0190"
                                                                    fncSmallBoreBaseCheck = False
                                                                    Exit Try
                                                                End If
                                                            Case "50"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 36 Then
                                                                    intKtbnStrcSeqNo = 8
                                                                    strMessageCd = "W0190"
                                                                    fncSmallBoreBaseCheck = False
                                                                    Exit Try
                                                                End If
                                                            Case "63"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 41 Then
                                                                    intKtbnStrcSeqNo = 8
                                                                    strMessageCd = "W0190"
                                                                    fncSmallBoreBaseCheck = False
                                                                    Exit Try
                                                                End If
                                                            Case "80"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 44 Then
                                                                    intKtbnStrcSeqNo = 8
                                                                    strMessageCd = "W0190"
                                                                    fncSmallBoreBaseCheck = False
                                                                    Exit Try
                                                                End If
                                                            Case "100"
                                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 50 Then
                                                                    intKtbnStrcSeqNo = 8
                                                                    strMessageCd = "W0190"
                                                                    fncSmallBoreBaseCheck = False
                                                                    Exit Try
                                                                End If
                                                        End Select
                                                End Select
                                            End If
                                    End Select
                                Case "T0H", "T5H"
                                    Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                                        Case "00", "LB", "FA", "FB", "FC", _
                                             "CA", "CB"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 20 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 15 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 15 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 15 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 20 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 25 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 25 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "T"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 60 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 65 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 70 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TC", "TF"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R", "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 110 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 135 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 110 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 115 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 125 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "T", "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 175 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 135 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 110 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 115 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 125 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TA", "TD", "TB", "TE"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 50 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 60 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 50 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 55 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 60 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                    End Select
                                Case "T0V", "T5V"
                                    Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                                        Case "00", "LB", "FA", "FB", "FC", _
                                             "CA", "CB"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 10 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 15 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 20 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 25 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "T"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 60 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 65 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 70 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TC", "TF"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R", "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 110 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 135 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 95 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 85 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 95 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "T", "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 145 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 135 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 100 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 105 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 115 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TA", "TD", "TB", "TE"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 50 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 60 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                    End Select
                                Case "T2H", "T3H"
                                    Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                                        Case "00", "LB", "FA", "FB", "FC", _
                                             "CA", "CB"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 20 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50", "63", "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 15 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 10 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 20 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50", "63", "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 15 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "T"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 25 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 30 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TC", "TF"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R", "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 105 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 110 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 115 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 125 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "T", "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 165 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 105 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 110 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 115 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 125 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TA", "TD", "TB", "TE"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 50 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 50 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 55 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 60 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                    End Select
                                Case "T2V", "T3V"
                                    Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                                        Case "00", "LB", "FA", "FB", "FC", _
                                             "CA", "CB"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 10 Then
                                                        intKtbnStrcSeqNo = 8
                                                        strMessageCd = "W0190"
                                                        fncSmallBoreBaseCheck = False
                                                        Exit Try
                                                    End If
                                                Case "D"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 15 Then
                                                        intKtbnStrcSeqNo = 8
                                                        strMessageCd = "W0190"
                                                        fncSmallBoreBaseCheck = False
                                                        Exit Try
                                                    End If
                                                Case "T"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 25 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 30 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TC", "TF"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R", "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 75 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 80 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 85 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 95 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "T", "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 135 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 75 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 85 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 90 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 100 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TA", "TD", "TB", "TE"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 35 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 30 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 35 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                    End Select
                                Case "T2YH", "T3YH", "T2JH", "T2YD", "T2YDT", "T2YDU", _
                                     "T2YLH", "T3YLH", "T1H", "T2WH", "T3WH"
                                    Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                                        Case "00", "LB", "FA", "FB", "FC", _
                                             "CA", "CB"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 20 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50", "63", "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 15 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 10 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 20 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50", "63", "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 15 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "T"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 25 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 30 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TC", "TF"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R", "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 105 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 100 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 105 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 110 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 120 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "T", "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 165 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 100 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 105 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 110 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 120 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TA", "TD", "TB", "TE"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 50 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 50 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 55 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 60 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                    End Select
                                Case "T2YV", "T3YV", "T2JV", "T2YLV", "T3YLV", "T1V", "T2WV", "T3WV"
                                    Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                                        Case "00", "LB", "FA", "FB", "FC", _
                                             "CA", "CB"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 10 Then
                                                        intKtbnStrcSeqNo = 8
                                                        strMessageCd = "W0190"
                                                        fncSmallBoreBaseCheck = False
                                                        Exit Try
                                                    End If
                                                Case "D"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 15 Then
                                                        intKtbnStrcSeqNo = 8
                                                        strMessageCd = "W0190"
                                                        fncSmallBoreBaseCheck = False
                                                        Exit Try
                                                    End If
                                                Case "T"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 25 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 30 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TC", "TF"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R", "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 75 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 70 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 75 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 80 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 90 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "T", "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 135 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 75 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 85 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 90 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 100 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TA", "TD", "TB", "TE"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 35 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 30 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 35 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                    End Select
                                Case "T8H"
                                    Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                                        Case "00", "LB", "FA", "FB", "FC", _
                                             "CA", "CB"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 15 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 10 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 15 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 20 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 25 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "T"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 60 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 65 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TC", "TF"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R", "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 95 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 115 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 95 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 100 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 110 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "T", "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 155 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 135 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 110 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 110 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 125 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TA", "TD", "TB", "TE"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 50 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 50 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 55 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                    End Select
                                Case "T8V"
                                    Select Case Trim(objKtbnStrc.strcSelection.strOpSymbol(4).Trim)
                                        Case "00", "LB", "FA", "FB", "FC", _
                                             "CA", "CB"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 10 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 15 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 20 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 25 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "T"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 45 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40", "50", "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 60 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80", "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 65 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TC", "TF"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R", "D"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 85 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 115 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 75 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 70 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 80 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                                Case "T", "4"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 125 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 135 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 110 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 115 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 125 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                        Case "TA", "TD", "TB", "TE"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                                Case "H", "R"
                                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                        Case "40"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "50"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 50 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "63"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 35 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "80"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 35 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                        Case "100"
                                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                                                intKtbnStrcSeqNo = 8
                                                                strMessageCd = "W0190"
                                                                fncSmallBoreBaseCheck = False
                                                                Exit Try
                                                            End If
                                                    End Select
                                            End Select
                                    End Select
                            End Select
                        Else
                            'スイッチが"E0"の時
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                '支持形式による分類
                                Case "00", "LB", "FA", "FB", "FC", "CA", "CB"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                        'スイッチ数による分類
                                        Case "H", "R", "D"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                '口径による分類
                                                Case "40"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 150 Then
                                                        intKtbnStrcSeqNo = 8
                                                        strMessageCd = "W0200"
                                                        fncSmallBoreBaseCheck = False
                                                        Exit Try
                                                    End If
                                                Case "50", "63", "80"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 145 Then
                                                        intKtbnStrcSeqNo = 8
                                                        strMessageCd = "W0200"
                                                        fncSmallBoreBaseCheck = False
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 140 Then
                                                        intKtbnStrcSeqNo = 8
                                                        strMessageCd = "W0200"
                                                        fncSmallBoreBaseCheck = False
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case "T"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 335 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncSmallBoreBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                Case "TC", "TF"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                        'スイッチ数による分類
                                        Case "H", "R", "D"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 335 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncSmallBoreBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "T"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 390 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncSmallBoreBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                Case "TA", "TD"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                        'スイッチ数による分類
                                        Case "H"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                '口径による分類
                                                Case "40"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 150 Then
                                                        intKtbnStrcSeqNo = 8
                                                        strMessageCd = "W0200"
                                                        fncSmallBoreBaseCheck = False
                                                        Exit Try
                                                    End If
                                                Case "50", "63", "80"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 145 Then
                                                        intKtbnStrcSeqNo = 8
                                                        strMessageCd = "W0200"
                                                        fncSmallBoreBaseCheck = False
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 140 Then
                                                        intKtbnStrcSeqNo = 8
                                                        strMessageCd = "W0200"
                                                        fncSmallBoreBaseCheck = False
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case Else
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0200"
                                            fncSmallBoreBaseCheck = False
                                            Exit Try
                                    End Select
                                Case "TB", "TE"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                        'スイッチ数による分類
                                        Case "R"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                                '口径による分類
                                                Case "40"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 150 Then
                                                        intKtbnStrcSeqNo = 8
                                                        strMessageCd = "W0200"
                                                        fncSmallBoreBaseCheck = False
                                                        Exit Try
                                                    End If
                                                Case "50", "63", "80"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 145 Then
                                                        intKtbnStrcSeqNo = 8
                                                        strMessageCd = "W0200"
                                                        fncSmallBoreBaseCheck = False
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 140 Then
                                                        intKtbnStrcSeqNo = 8
                                                        strMessageCd = "W0200"
                                                        fncSmallBoreBaseCheck = False
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case Else
                                            intKtbnStrcSeqNo = 12
                                            strMessageCd = "W0740"
                                            fncSmallBoreBaseCheck = False
                                            Exit Try
                                    End Select
                            End Select
                        End If
                    End If
                Case Else
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 1 Then
                        intKtbnStrcSeqNo = 8
                        strMessageCd = "W0200"
                        fncSmallBoreBaseCheck = False
                        Exit Try
                    End If
            End Select

            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                '最大値チェック
                'バリエーションによる分類
                Case "", "V", "H", "T", "T2", "G", "G1", "VG", "VG1", "TG1", "T2G1"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        '口径による分類
                        Case "40"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) > 1600 Then
                                intKtbnStrcSeqNo = 8
                                strMessageCd = "W0200"
                                fncSmallBoreBaseCheck = False
                                Exit Try
                            End If
                        Case "50"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) > 2000 Then
                                intKtbnStrcSeqNo = 8
                                strMessageCd = "W0200"
                                fncSmallBoreBaseCheck = False
                                Exit Try
                            End If
                        Case "63", "80", "100"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) > 2500 Then
                                intKtbnStrcSeqNo = 8
                                strMessageCd = "W0200"
                                fncSmallBoreBaseCheck = False
                                Exit Try
                            End If
                    End Select
                Case Else
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        '口径による分類
                        Case "40"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) > 1600 Then
                                intKtbnStrcSeqNo = 8
                                strMessageCd = "W0200"
                                fncSmallBoreBaseCheck = False
                                Exit Try
                            End If
                        Case "50", "63", "80", "100"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) > 1900 Then
                                intKtbnStrcSeqNo = 8
                                strMessageCd = "W0200"
                                fncSmallBoreBaseCheck = False
                                Exit Try
                            End If
                    End Select
            End Select

            '付属品チェック
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(14), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "B1", "B3"
                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "CB" And _
                           objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("Y") < 0 Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0290"
                            fncSmallBoreBaseCheck = False
                            Exit Try
                        End If
                    Case "B2"
                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "CA" And _
                           objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("I") < 0 Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0300"
                            fncSmallBoreBaseCheck = False
                            Exit Try
                        End If
                End Select
            Next

            'ロッド先端特注
            If objKtbnStrc.strcSelection.strRodEndOption.Trim <> "" Then
                '付属品「I」「Y」選択時はロッド先端は"N13","N15"以外選択不可
                If objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("I") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("Y") >= 0 Then
                    If objKtbnStrc.strcSelection.strRodEndOption.IndexOf("N13") >= 0 Or _
                       objKtbnStrc.strcSelection.strRodEndOption.IndexOf("N15") >= 0 Then
                    Else
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0350"
                        fncSmallBoreBaseCheck = False
                        Exit Try
                    End If
                End If
            End If

            'オプション外
            If objKtbnStrc.strcSelection.strOtherOption.Trim <> "" Then
                '支持金具90°回転(K1)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K1") >= 0 Then
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "00", "FA", "FC"
                            intKtbnStrcSeqNo = 4
                            strMessageCd = "W0430"
                            fncSmallBoreBaseCheck = False
                            Exit Try
                    End Select
                End If

                '支持金具180°回転(K2)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K2") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "LB" Then
                        intKtbnStrcSeqNo = 4
                        strMessageCd = "W0440"
                        fncSmallBoreBaseCheck = False
                        Exit Try
                    End If
                End If

                '支持金具270°回転(K3)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K3") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "LB" Then
                        intKtbnStrcSeqNo = 4
                        strMessageCd = "W0450"
                        fncSmallBoreBaseCheck = False
                        Exit Try
                    End If
                End If

                'トラニオン位置
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("AQ") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "TC" And _
                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "TF" Then
                        intKtbnStrcSeqNo = 4
                        strMessageCd = "W0460"
                        fncSmallBoreBaseCheck = False
                        Exit Try
                    End If
                End If

                'P5
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P5") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "CB" And _
                       objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("Y") < 0 Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0470"
                        fncSmallBoreBaseCheck = False
                        Exit Try
                    End If
                End If

                'P7
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P7") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("I") < 0 And _
                       objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("Y") < 0 Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0480"
                        fncSmallBoreBaseCheck = False
                        Exit Try
                    End If
                End If

                'P8
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P8") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("Y") < 0 Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0490"
                        fncSmallBoreBaseCheck = False
                        Exit Try
                    End If
                End If

                'MM()H*
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("MM") >= 0 Then
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "FB", "CA", "CB"
                            intKtbnStrcSeqNo = 4
                            strMessageCd = "W0750"
                            fncSmallBoreBaseCheck = False
                            Exit Try
                    End Select
                End If

                'M1
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("M1") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(2).IndexOf("K") >= 0 Then
                        intKtbnStrcSeqNo = 2
                        strMessageCd = "W0760"
                        fncSmallBoreBaseCheck = False
                        Exit Try
                    End If
                End If

                'J9
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("J9") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("J") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(2).IndexOf("G") >= 0 Then
                        intKtbnStrcSeqNo = 13
                        strMessageCd = "W0770"
                        fncSmallBoreBaseCheck = False
                        Exit Try
                    End If
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncBigBoreBaseCheck
    '*【処理】
    '*  φ125～180ベースチェック
    '*【概要】
    '*  φ125～180ベースをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Public Function fncBigBoreBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncBigBoreBaseCheck = True

            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                'バリエーションによる分類
                Case "LN", "LH", "LNG", "LNG1", "LHG", "LHG1"
                    If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "" Then
                        '支持形式判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "LB", "FA", "FB", "CA", "CB"
                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 1 Then
                                    intKtbnStrcSeqNo = 8
                                    strMessageCd = "W0200"
                                    fncBigBoreBaseCheck = False
                                    Exit Try
                                End If
                            Case Else
                                '口径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "125"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 30 Then
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0200"
                                            fncBigBoreBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "140"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 32 Then
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0200"
                                            fncBigBoreBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "160"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 34 Then
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0200"
                                            fncBigBoreBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "180"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 35 Then
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0200"
                                            fncBigBoreBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                        End Select
                    Else
                        '支持形式判定
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                            Case "LB", "FA", "FB", "FC", "CA", "CB"
                                'スイッチ数判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                    Case "H", "D", "R"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 20 Then
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0200"
                                            fncBigBoreBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "T"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0200"
                                            fncBigBoreBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "4"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 55 Then
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0200"
                                            fncBigBoreBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            Case "TC", "TF"
                                '口径判定
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                    Case "125"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 120 Then
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0200"
                                            fncBigBoreBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "140"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 125 Then
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0200"
                                            fncBigBoreBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "160"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 130 Then
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0200"
                                            fncBigBoreBaseCheck = False
                                            Exit Try
                                        End If
                                    Case "180"
                                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 135 Then
                                            intKtbnStrcSeqNo = 8
                                            strMessageCd = "W0200"
                                            fncBigBoreBaseCheck = False
                                            Exit Try
                                        End If
                                End Select
                            Case "TA", "TD"
                                If Len(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) <> 0 Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(12).Trim <> "H" Then
                                        intKtbnStrcSeqNo = 8
                                        strMessageCd = "W0200"
                                        fncBigBoreBaseCheck = False
                                        Exit Try
                                    Else
                                        '口径判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            Case "125"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 70 Then
                                                    intKtbnStrcSeqNo = 8
                                                    strMessageCd = "W0200"
                                                    fncBigBoreBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "140"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 75 Then
                                                    intKtbnStrcSeqNo = 8
                                                    strMessageCd = "W0200"
                                                    fncBigBoreBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "160"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 80 Then
                                                    intKtbnStrcSeqNo = 8
                                                    strMessageCd = "W0200"
                                                    fncBigBoreBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "180"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 85 Then
                                                    intKtbnStrcSeqNo = 8
                                                    strMessageCd = "W0200"
                                                    fncBigBoreBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                    End If
                                End If
                            Case "TB", "TE"
                                If Len(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) <> 0 Then
                                    If objKtbnStrc.strcSelection.strOpSymbol(12).Trim <> "R" Then
                                        intKtbnStrcSeqNo = 12
                                        strMessageCd = "W0740"
                                        fncBigBoreBaseCheck = False
                                        Exit Try
                                    Else
                                        '口径判定
                                        Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                            Case "125"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 70 Then
                                                    intKtbnStrcSeqNo = 8
                                                    strMessageCd = "W0200"
                                                    fncBigBoreBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "140"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 75 Then
                                                    intKtbnStrcSeqNo = 8
                                                    strMessageCd = "W0200"
                                                    fncBigBoreBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "160"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 80 Then
                                                    intKtbnStrcSeqNo = 8
                                                    strMessageCd = "W0200"
                                                    fncBigBoreBaseCheck = False
                                                    Exit Try
                                                End If
                                            Case "180"
                                                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 85 Then
                                                    intKtbnStrcSeqNo = 8
                                                    strMessageCd = "W0200"
                                                    fncBigBoreBaseCheck = False
                                                    Exit Try
                                                End If
                                        End Select
                                    End If
                                End If
                        End Select
                    End If
                Case Else
                    '支持形式判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "LB", "FA", "FB", "FC", "CA", "CB"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 1 Then
                                intKtbnStrcSeqNo = 8
                                strMessageCd = "W0200"
                                fncBigBoreBaseCheck = False
                                Exit Try
                            End If
                        Case "TA", "TB", "TC", "TD", "TE", "TF"
                            '口径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "125"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 30 Then
                                        intKtbnStrcSeqNo = 8
                                        strMessageCd = "W0200"
                                        fncBigBoreBaseCheck = False
                                        Exit Try
                                    End If
                                Case "140"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 32 Then
                                        intKtbnStrcSeqNo = 8
                                        strMessageCd = "W0200"
                                        fncBigBoreBaseCheck = False
                                        Exit Try
                                    End If
                                Case "160"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 34 Then
                                        intKtbnStrcSeqNo = 8
                                        strMessageCd = "W0200"
                                        fncBigBoreBaseCheck = False
                                        Exit Try
                                    End If
                                Case "180"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 35 Then
                                        intKtbnStrcSeqNo = 8
                                        strMessageCd = "W0200"
                                        fncBigBoreBaseCheck = False
                                        Exit Try
                                    End If
                            End Select
                    End Select
            End Select

            '最大値チェック
            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) > 2000 Then
                If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 35 Then
                    intKtbnStrcSeqNo = 8
                    strMessageCd = "W0200"
                    fncBigBoreBaseCheck = False
                    Exit Try
                End If
            End If

            ' 付属品をチェック
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(14), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "B1"
                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "CB" And _
                           objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("Y") < 0 Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0290"
                            fncBigBoreBaseCheck = False
                            Exit Try
                        End If
                    Case "B2"
                        If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "CA" And _
                           objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("I") < 0 Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0300"
                            fncBigBoreBaseCheck = False
                            Exit Try
                        End If
                End Select
            Next

            'ロッド先端特注
            If objKtbnStrc.strcSelection.strRodEndOption.Trim <> "" Then
                '付属品「I」「Y」選択時はロッド先端は選択不可
                If objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("I") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("Y") >= 0 Then
                    If objKtbnStrc.strcSelection.strRodEndOption.IndexOf("N13") >= 0 Or _
                       objKtbnStrc.strcSelection.strRodEndOption.IndexOf("N15") >= 0 Then
                    Else
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0350"
                        fncBigBoreBaseCheck = False
                        Exit Try
                    End If
                End If
            End If

            'オプション外
            If objKtbnStrc.strcSelection.strOtherOption.Trim <> "" Then
                '支持金具90°回転(K1)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K1") >= 0 Then
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "LB", "FB", "CA", "CB", "TA", _
                             "TB", "TC", "TD", "TE", "TF"
                        Case Else
                            intKtbnStrcSeqNo = 4
                            strMessageCd = "W0430"
                            fncBigBoreBaseCheck = False
                            Exit Try
                    End Select
                End If

                '支持金具180°回転(K2)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K2") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "LB" Then
                        intKtbnStrcSeqNo = 4
                        strMessageCd = "W0440"
                        fncBigBoreBaseCheck = False
                        Exit Try
                    End If
                End If

                '支持金具270°回転(K3)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K3") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "LB" Then
                        intKtbnStrcSeqNo = 4
                        strMessageCd = "W0450"
                        fncBigBoreBaseCheck = False
                        Exit Try
                    End If
                End If

                'トラニオン位置
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("AQ") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "TC" And _
                       objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "TF" Then
                        intKtbnStrcSeqNo = 4
                        strMessageCd = "W0460"
                        fncBigBoreBaseCheck = False
                        Exit Try
                    End If
                End If

                'P5
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P5") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(4).Trim <> "CB" And _
                       objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("Y") < 0 Then
                        intKtbnStrcSeqNo = 4
                        strMessageCd = "W0470"
                        fncBigBoreBaseCheck = False
                        Exit Try
                    End If
                End If

                'P7
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P7") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("I") < 0 And _
                       objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("Y") < 0 Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0480"
                        fncBigBoreBaseCheck = False
                        Exit Try
                    End If
                End If

                'P8
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P8") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("Y") < 0 Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0490"
                        fncBigBoreBaseCheck = False
                        Exit Try
                    End If
                End If

                'MX()H*
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("MX") >= 0 Then
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "FB", "CA", "CB"
                            intKtbnStrcSeqNo = 4
                            strMessageCd = "W0750"
                            fncBigBoreBaseCheck = False
                            Exit Try
                    End Select
                End If

                'J9
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("J9") >= 0 Then
                    If (objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("J") >= 0 Or _
                        objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("K") >= 0 Or _
                        objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("L") >= 0) Or _
                        objKtbnStrc.strcSelection.strOpSymbol(2).IndexOf("G") >= 0 Then
                        intKtbnStrcSeqNo = 13
                        strMessageCd = "W0780"
                        fncBigBoreBaseCheck = False
                        Exit Try
                    End If
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncStandardBaseCheck
    '*【処理】
    '*  ＪＳＣ４ φ125～180ベースチェック
    '*【概要】
    '*  ＪＳＣ４ φ125～180ベースをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Public Function fncStandardBoreBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean


        Try

            fncStandardBoreBaseCheck = True

            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                'バリエーションによる分類
                Case "LN", "LH"
                    '支持形式判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "LB", "FA", "FB", "CA", "CB"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                'スイッチ数による分類
                                Case "H", "R", "D"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 20 Then
                                        intKtbnStrcSeqNo = 8
                                        strMessageCd = "W0200"
                                        fncStandardBoreBaseCheck = False
                                        Exit Try
                                    End If
                                Case "T"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 40 Then
                                        intKtbnStrcSeqNo = 8
                                        strMessageCd = "W0200"
                                        fncStandardBoreBaseCheck = False
                                        Exit Try
                                    End If
                                Case "4"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 55 Then
                                        intKtbnStrcSeqNo = 8
                                        strMessageCd = "W0200"
                                        fncStandardBoreBaseCheck = False
                                        Exit Try
                                    End If
                            End Select
                        Case "TC"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                'スイッチ数による分類
                                Case "H", "R", "D", "T", "4"
                                    '口径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "125"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 120 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "140"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 125 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "160"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 130 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "180"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 135 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                Case Else
                                    '口径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "125"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 30 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "140"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 32 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "160"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 34 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "180"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 35 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                            End Select
                        Case "TA", "TB"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(12).Trim
                                'スイッチ数による分類
                                Case "H", "R", "D", "T", "4"
                                    '口径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "125"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 70 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "140"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 75 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "160"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 80 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "180"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 85 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                                Case Else
                                    '口径判定
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                        Case "125"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 30 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "140"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 32 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "160"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 34 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                        Case "180"
                                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 35 Then
                                                intKtbnStrcSeqNo = 8
                                                strMessageCd = "W0200"
                                                fncStandardBoreBaseCheck = False
                                                Exit Try
                                            End If
                                    End Select
                            End Select
                    End Select
                Case "N", "H", "T"
                    '支持形式判定
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "TA", "TB", "TC"
                            '口径判定
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                                Case "125"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 30 Then
                                        intKtbnStrcSeqNo = 8
                                        strMessageCd = "W0200"
                                        fncStandardBoreBaseCheck = False
                                        Exit Try
                                    End If
                                Case "140"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 32 Then
                                        intKtbnStrcSeqNo = 8
                                        strMessageCd = "W0200"
                                        fncStandardBoreBaseCheck = False
                                        Exit Try
                                    End If
                                Case "160"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 34 Then
                                        intKtbnStrcSeqNo = 8
                                        strMessageCd = "W0200"
                                        fncStandardBoreBaseCheck = False
                                        Exit Try
                                    End If
                                Case "180"
                                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(8).Trim) < 35 Then
                                        intKtbnStrcSeqNo = 8
                                        strMessageCd = "W0200"
                                        fncStandardBoreBaseCheck = False
                                        Exit Try
                                    End If
                            End Select
                    End Select
            End Select


        Catch ex As Exception

            Throw ex

        End Try

    End Function

End Module
