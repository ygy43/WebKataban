'************************************************************************************
'*  ProgramID  ：KHCylinderSSD2Check
'*  Program名  ：シリンダＳＳＤ２シリーズチェックモジュール
'*
'*                                      作成日：2008/01/11   作成者：NII A.Takahashi
'*
'*  概要       ：ＳＳＤ２／ＳＳＤ２ーＫ
'*  ・受付No：RM0906034  二次電池対応機器対応
'*                                      更新日：2009/08/05   更新者：Y.Miura
'************************************************************************************
Module KHCylinderSSD2Check

    '********************************************************************************************
    '*【関数名】
    '*  fncCheckSelectOption
    '*【処理】
    '*  シリンダチェック
    '*【概要】
    '*  シリンダＳＳＤ２シリーズをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*【修正履歴】
    '*                                      更新日：2008/05/07   更新者：T.Sato
    '*  ・受付No：RM0802088対応　バリエーション（'Ｄ','Ｍ','Ｑ','Ｘ','Ｙ'）追加に伴う修正
    '********************************************************************************************
    Public Function fncCheckSelectOption(ByVal objKtbnStrc As KHKtbnStrc, _
                                         ByRef intKtbnStrcSeqNo As Integer, _
                                         ByRef strOptionSymbol As String, _
                                         ByRef strMessageCd As String) As Boolean

        Try

            fncCheckSelectOption = True

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "SSD2"
                    '基本ベース毎にチェック
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        '2010/10/05 MOD RM1010017(11月VerUP:SSD2シリーズ) START--->
                        Case ""
                            '基本ベースチェック
                            If fncStandardBaseCheck(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case "4"
                            'RM0906034 2009/08/05 Y.Miura　二次電池対応機種追加
                            'Case "", "4"
                            ''Case ""
                            '2010/10/05 MOD RM1010017(11月VerUP:SSD2シリーズ) <---END
                            '２次電池チェック
                            If fncP4BaseCheck(objKtbnStrc, _
                                              intKtbnStrcSeqNo, _
                                              strOptionSymbol, _
                                              strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case "D", "E"
                            '両ロッドベースチェック
                            If fncDoubleRodBaseCheck(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If

                            '2010/11/02 MOD RM1011020(12月VerUP:SSD2シリーズ) START--->
                        Case "K"
                            '高荷重ベースチェック
                            If fncHighLoadBaseCheck(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If

                        Case "L"
                            '高荷重ベースチェック
                            If fncHighLoadBaseP4Check(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If

                        Case "6", "E"
                            'ロングストローク（両ロッド）ベース(Ｐ４)チェック
                            If fncLongBaseP4Check(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                            ''RM0906034 2009/08/05 Y.Miura　二次電池対応機種追加
                            ''Case "K"
                            'Case "K", "L"
                            '    '高荷重ベースチェック
                            '    If fncHighLoadBaseCheck(objKtbnStrc, _
                            '                            intKtbnStrcSeqNo, _
                            '                            strOptionSymbol, _
                            '                            strMessageCd) = False Then
                            '        fncCheckSelectOption = False
                            '    End If
                            'Case "M"
                            '    '回り止めベースチェック
                            '    If fncNonRotatingBaseCheck(objKtbnStrc, _
                            '                            intKtbnStrcSeqNo, _
                            '                            strOptionSymbol, _
                            '                            strMessageCd) = False Then
                            '        fncCheckSelectOption = False
                            '    End If
                            'Case "Q"
                            '    '落下防止ベースチェック
                            '    If fncPositionLockingBaseCheck(objKtbnStrc, _
                            '                            intKtbnStrcSeqNo, _
                            '                            strOptionSymbol, _
                            '                            strMessageCd) = False Then
                            '        fncCheckSelectOption = False
                            '    End If
                            'Case "X"
                            '    '押出しベースチェック
                            '    If fncSpringReturnBaseCheck(objKtbnStrc, _
                            '                            intKtbnStrcSeqNo, _
                            '                            strOptionSymbol, _
                            '                            strMessageCd) = False Then
                            '        fncCheckSelectOption = False
                            '    End If
                            'Case "Y"
                            '    '引込みベースチェック
                            '    If fncSpringExtendBaseCheck(objKtbnStrc, _
                            '                            intKtbnStrcSeqNo, _
                            '                            strOptionSymbol, _
                            '                            strMessageCd) = False Then
                            '        fncCheckSelectOption = False
                            '    End If
                            '2010/11/02 MOD RM1011020(12月VerUP:SSD2シリーズ) <---END

                    End Select
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) START--->
    '********************************************************************************************
    '*【関数名】
    '*  fncStandardBaseCheck
    '*【処理】
    '*  基本ベースチェック
    '*【概要】
    '*  基本ベースをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncStandardBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Try

            fncStandardBaseCheck = True

            '*-----オプションチェック-----*
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1)
                Case "G", "G2", "G3"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4)
                        Case "16", "20", "25", "32"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(20)
                                Case "FA", "LB"
                                    intKtbnStrcSeqNo = 20
                                    strMessageCd = "W9050"
                                    fncStandardBaseCheck = False
                                    Exit Try
                            End Select
                    End Select
            End Select


            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            Dim selList As New ArrayList

            Select Case objKtbnStrc.strcSelection.strOpSymbol(1)
                Case "T1L"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(18)
                        Case "R", "H"
                            selList.Add("10:16,32,40,50,63,80,100")
                            selList.Add("15:20,25")

                        Case "D"
                            selList.Add("20:16,25,32,40,50,63,80,100")
                            selList.Add("25:20")

                    End Select

                    'ストロークのチェック
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 2) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If

                    '2010/11/02 ADD RM1011020(12月VerUP:SSD2シリーズ) START--->
                    selList.Clear()
                    selList.Add("50:20,25")

                    'ストロークのチェック
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                Case ""

                    'バリエーション②
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "L4" Then
                        selList.Add("20:*")

                        'ストロークのチェック
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 2) = False Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If

                    End If

                    selList.Clear()
                    selList.Add("50:20,25")

                    'ストロークのチェック
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If

                Case "G1"
                    'バリエーション②
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "L4" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(18).Trim
                            Case "R", "H", "D"
                                selList.Add("20:*")
                            Case "T"
                                selList.Add("35:*")
                        End Select

                        'ストロークのチェック
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 2) = False Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If

                    End If

                    selList.Clear()
                    selList.Add("50:20,25")

                    'ストロークのチェック
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If

                Case "T1", "O", "G", "G2", "G3", "G4", "G5"
                    selList.Add("50:20,25")

                    'ストロークのチェック
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                    '2010/11/02 ADD RM1011020(12月VerUP:SSD2シリーズ) <---END

                Case "W"
                    selList.Add("30:12,16")
                    selList.Add("50:20,25,32,40,50,63,80,100")
                    selList.Add("300:125,140,160")

                    'Ｓ１ストロークの有無
                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> 0 Then
                        'Ｓ１ストロークチェック
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 1) = False Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If

                    'Ｓ２ストロークの有無
                    If objKtbnStrc.strcSelection.strOpSymbol(14) <> 0 Then

                        'Ｓ２のチェック
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 1) = False Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If
                    '2010/11/02 ADD RM1011020(12月VerUP:SSD2シリーズ) START--->
                Case "M"

                    selList.Add("5:12,16,20,25,32,40")
                    selList.Add("10:50,63")

                    'ストロークのチェック
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 2) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If

                Case "X", "Y"
                    'Ｓ２ストロークチェック
                    If objKtbnStrc.strcSelection.strOpSymbol(14) <> 0 Then
                        '口径別
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4)
                            Case "12", "16", "20", "25", "32", "40"
                                'Ｓ２チェック
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "5", "10"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncStandardBaseCheck = False
                                        Exit Try
                                End Select
                            Case "50"
                                'Ｓ２チェック
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "10", "20"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncStandardBaseCheck = False
                                        Exit Try
                                End Select
                        End Select
                    End If
                Case "Q"
                    'Ｓ２ストロークチェック
                    If objKtbnStrc.strcSelection.strOpSymbol(14) <> 0 Then
                        '口径別
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4)
                            Case "20", "25", "32", "40", "50", "63"
                                'Ｓ２チェック
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "10", "15", "20", "25", "50", "75", "100"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncStandardBaseCheck = False
                                        Exit Try
                                End Select
                            Case "80", "100"
                                'Ｓ２チェック
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "25", "50", "75", "100"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncStandardBaseCheck = False
                                        Exit Try
                                End Select
                        End Select
                    End If
                    '2010/11/02 ADD RM1011020(12月VerUP:SSD2シリーズ) <---END
                Case Else
                    'チェックなし
            End Select

            'Ｓ１スイッチ毎のチェック
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(11).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 7
                strMessageCd = "W0200"
                fncStandardBaseCheck = False
                Exit Try
            End If

            'Ｓ２スイッチ毎のチェック
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(16).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(18).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 14
                strMessageCd = "W0200"
                fncStandardBaseCheck = False
                Exit Try
            End If

            '*-----<< Ⅱ．最大ストロークとバリエーションのチェック >>-----*
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "M"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "20", "25"
                            If CDec(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 50 Then
                                intKtbnStrcSeqNo = 14
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        Case Else
                    End Select
                Case Else
            End Select

            '*-----<< Ⅱ．最大ストロークとゴムクッションのチェック >>-----*
            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(5).Trim, "D") = 0 Then
            Else
                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    Case "12", "16"
                        If CDec(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 30 Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    Case "20", "25"
                        If CDec(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 50 Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    Case "32", "40", "50", "63", "80", "100"
                        If CDec(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 100 Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                End Select
            End If

            '*-----<< Ⅲ．最大ストロークと最小ストロークの相関チェック >>-----*
            '二段形の時
            If objKtbnStrc.strcSelection.strOpSymbol(1) = "W" Then

                If IsNumeric(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) AndAlso _
                IsNumeric(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) AndAlso _
                CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) Then
                    intKtbnStrcSeqNo = 14
                    strMessageCd = "W0610"
                    fncStandardBaseCheck = False
                    Exit Try

                End If
            End If

            '201012/10 ADD RM1012055(1月VerUP::SSD2シリーズ) START--->
            '*-----<< オプション「中間ストローク専用本体」チェック >>-----*
            Dim strOp() As String
            strOp = Split(objKtbnStrc.strcSelection.strOpSymbol(19).Trim, ",")
            If Not fncOptionSCheck(strOp, _
                                objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                objKtbnStrc.strcSelection.strOpSymbol(14).Trim) Then
                intKtbnStrcSeqNo = 14
                strMessageCd = "W0830"
                fncStandardBaseCheck = False
                Exit Try
            End If
            '2010/12/10 ADD RM1012055(1月VerUP:SSD2シリーズ) <---END

        Catch ex As Exception

            Throw ex

        End Try

    End Function
    '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) <---END

    '********************************************************************************************
    '*【関数名】
    '*  fncP4BaseCheck
    '*【処理】
    '*  二次電池ベースチェック
    '*【概要】
    '*  二次電池ベースをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    '2010/10/05 MOD RM1010017(11月VerUP:SSD2シリーズ) START--->
    Private Function fncP4BaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean
        'Private Function fncStandardBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
        '2010/10/05 MOD RM1010017(11月VerUP:SSD2シリーズ) <---END

        Try

            fncP4BaseCheck = True

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            'スイッチ毎のチェック
            'RM1210067 2013/02/01 Y.Tachi ローカル版との差異修正
            'Ｓ１
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(11).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 7
                strMessageCd = "W0200"
                fncP4BaseCheck = False
                Exit Try
            End If
            'Ｓ２
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(16).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(18).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 14
                strMessageCd = "W0200"
                fncP4BaseCheck = False
                Exit Try
            End If
            '*-----<< Ⅱ．最大ストロークとゴムクッションのチェック >>-----*

            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(4).Trim, "D") = 0 Then
            Else
                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    Case "12", "16"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 30 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W0200"
                            fncP4BaseCheck = False
                            Exit Try
                        End If
                    Case "20", "25"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 50 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W0200"
                            fncP4BaseCheck = False
                            Exit Try
                        End If
                    Case "32", "40", "50", "63", "80", "100"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 100 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W0200"
                            fncP4BaseCheck = False
                            Exit Try
                        End If
                End Select
            End If

            Dim selList As New ArrayList

            Select Case objKtbnStrc.strcSelection.strOpSymbol(1)
                Case "T1L"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(18)
                        Case "R", "H"
                            selList.Add("10:16,32,40,50,63,80,100")
                            selList.Add("15:20,25")

                        Case "D"
                            selList.Add("20:16,25,32,40,50,63,80,100")
                            selList.Add("25:20")

                    End Select

                    'ストロークのチェック
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 2) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncP4BaseCheck = False
                        Exit Try
                    End If

                    '2010/11/02 ADD RM1011020(12月VerUP:SSD2シリーズ) START--->
                    selList.Clear()
                    selList.Add("50:20,25")

                    'ストロークのチェック
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncP4BaseCheck = False
                        Exit Try
                    End If
                Case ""

                    'バリエーション②
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "L4" Then
                        selList.Add("20:*")

                        'ストロークのチェック
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 2) = False Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncP4BaseCheck = False
                            Exit Try
                        End If

                    End If

                    selList.Clear()
                    selList.Add("50:20,25")

                    'ストロークのチェック
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncP4BaseCheck = False
                        Exit Try
                    End If

                Case "G1"
                    'バリエーション②
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "L4" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(18).Trim
                            Case "R", "H", "D"
                                selList.Add("20:*")
                            Case "T"
                                selList.Add("35:*")
                        End Select

                        'ストロークのチェック
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 2) = False Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncP4BaseCheck = False
                            Exit Try
                        End If

                    End If

                    selList.Clear()
                    selList.Add("50:20,25")

                    'ストロークのチェック
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncP4BaseCheck = False
                        Exit Try
                    End If

                Case "T1", "O", "G", "G2", "G3", "G4", "G5"
                    selList.Add("50:20,25")

                    'ストロークのチェック
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncP4BaseCheck = False
                        Exit Try
                    End If
                    '2010/11/02 ADD RM1011020(12月VerUP:SSD2シリーズ) <---END

                Case "W"
                    selList.Add("30:12,16")
                    selList.Add("50:20,25,32,40,50,63,80,100")
                    selList.Add("300:125,140,160")

                    'Ｓ１ストロークの有無
                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> 0 Then
                        'Ｓ１ストロークチェック
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 1) = False Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncP4BaseCheck = False
                            Exit Try
                        End If
                    End If

                    'Ｓ２ストロークの有無
                    If objKtbnStrc.strcSelection.strOpSymbol(14) <> 0 Then

                        'Ｓ２のチェック
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 1) = False Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncP4BaseCheck = False
                            Exit Try
                        End If
                    End If
                    '2010/11/02 ADD RM1011020(12月VerUP:SSD2シリーズ) START--->
                Case "M"

                    selList.Add("5:12,16,20,25,32,40")
                    selList.Add("10:50,63")

                    'ストロークのチェック
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 2) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncP4BaseCheck = False
                        Exit Try
                    End If

                Case "X", "Y"
                    'Ｓ２ストロークチェック
                    If objKtbnStrc.strcSelection.strOpSymbol(14) <> 0 Then
                        '口径別
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4)
                            Case "12", "16", "20", "25", "32", "40"
                                'Ｓ２チェック
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "5", "10"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncP4BaseCheck = False
                                        Exit Try
                                End Select
                            Case "50"
                                'Ｓ２チェック
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "10", "20"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncP4BaseCheck = False
                                        Exit Try
                                End Select
                        End Select
                    End If
                Case "Q"
                    'Ｓ２ストロークチェック
                    If objKtbnStrc.strcSelection.strOpSymbol(14) <> 0 Then
                        '口径別
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4)
                            Case "20", "25", "32", "40", "50", "63"
                                'Ｓ２チェック
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "10", "15", "20", "25", "50", "75", "100"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncP4BaseCheck = False
                                        Exit Try
                                End Select
                            Case "80", "100"
                                'Ｓ２チェック
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "25", "50", "75", "100"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncP4BaseCheck = False
                                        Exit Try
                                End Select
                        End Select
                    End If
                Case Else
                    'チェックなし
            End Select

            '2010/10/05 DEL RM1010017(11月VerUP:SSD2シリーズ) START--->
            ''RM0906034 2009/08/05 Y.Miura　二次電池対応機種追加
            'If objKtbnStrc.strcSelection.strKeyKataban.Equals("4") Then
            '2010/10/05 DEL RM1010017(11月VerUP:SSD2シリーズ) <---END

            '基本ベースチェック　二次電池対応
            '↓2012/01/05 RM1201XXX intOptionPos(9→19)変更
            If fncP4Check(objKtbnStrc, _
                                    intKtbnStrcSeqNo, _
                                    strOptionSymbol, _
                                    strMessageCd, _
                                    19) = False Then
                fncP4BaseCheck = False
                Exit Try
            End If
            'End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncDoubleRodBaseCheck
    '*【処理】
    '*  両ロッドベースチェック
    '*【概要】
    '*  両ロッドベースをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncDoubleRodBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Try

            fncDoubleRodBaseCheck = True

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            '2010/11/02 ADD RM1011020(12月VerUP:SSD2シリーズ) START--->
            Dim selList As New ArrayList

            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "D"
                    selList.Add("5:25,32,40")
                    selList.Add("10:50,63,80,100")

                    'ストロークチェック
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                              selList, 2) = False Then
                        intKtbnStrcSeqNo = 5
                        strMessageCd = "W0200"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
            End Select
            '2010/11/02 ADD RM1011020(12月VerUP:SSD2シリーズ) <---END

            'スイッチ毎のチェック
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 5
                strMessageCd = "W0200"
                fncDoubleRodBaseCheck = False
                Exit Try
            End If

            '*-----<< Ⅱ．中間ストロークチェック >>-----*

            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Case "32", "40", "50", "63", "80", "100"
                    ' 中間ストロークチェック(5mm毎)
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) Mod 5 <> 0 Then
                        intKtbnStrcSeqNo = 5
                        strMessageCd = "W0510"
                        fncDoubleRodBaseCheck = False
                    End If

            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncHighLoadBaseCheck
    '*【処理】
    '*  高荷重ベースチェック
    '*【概要】
    '*  高荷重ベースをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncHighLoadBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Try

            fncHighLoadBaseCheck = True

            '*-----オプションチェック-----*
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1)
                Case "KG2", "KG3"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4)
                        Case "16", "20", "25", "32"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(20)
                                Case "FA", "LB"
                                    intKtbnStrcSeqNo = 20
                                    strMessageCd = "W9050"
                                    fncHighLoadBaseCheck = False
                                    Exit Try
                            End Select
                    End Select
            End Select

            '2010/11/02 ADD RM1011020(12月VerUP:SSD2シリーズ) START--->
            '*-----<< Ⅰ．ストロークチェック >>-----*
            Dim selListMin As New ArrayList
            'バリエーション①
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "K", "KG1"
                    '配管ねじ、クッション
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case "C", "GC", "NC"
                            '口径
                            selListMin.Add("5:20,25,32,40,50")
                            selListMin.Add("10:63,80,100")

                            If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                                      objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                      selListMin, 2) = False Then
                                intKtbnStrcSeqNo = 14
                                strMessageCd = "W0200"
                                fncHighLoadBaseCheck = False
                                Exit Try
                            End If

                            selListMin.Clear()

                    End Select
            End Select

            'バリエーション②
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "L"

                    'スイッチ
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                        Case ""
                            'Do Nothing
                        Case "T0H", "T0V", "T5H", "T5V"
                            '最小チェック値設定
                            selListMin.Add("10:12,16")
                            selListMin.Add("5:20,25,32,40,50,63,80,100")
                        Case "F2V", "F3V", "F2YV", "F3YV"
                            '最小チェック値設定
                            selListMin.Add("10:20")
                            selListMin.Add("5:12,16,25,32,40,50,63,80,100")
                        Case "T2H", "T2V", "T3H", "T3V", "T3PH", "T3PV", "F2H", "F2V", "F2YH", "F3YH"
                            '最小チェック値設定
                            selListMin.Add("5:*")
                        Case Else
                            '最小チェック値設定
                            selListMin.Add("10:*")
                    End Select

                    'バリエーション①
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "K", "KG1"
                            'スイッチ
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                                Case "T0H", "T0V", "T5H", "T5V", "T2H", "T2V", "T3H", "T3V"
                                    'Ｓ２：数
                                    If objKtbnStrc.strcSelection.strOpSymbol(18).Trim = "D" Then
                                        '最小チェック値設定
                                        selListMin.Clear()  '※上記で設定されている場合、最小チェック値を塗替える
                                        selListMin.Add("5:*")
                                    End If
                            End Select
                    End Select
                    '↓RM1212080 2012/12/05 Y.Tachi
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "K"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "12", "16"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                                        Case "T0H", "T0V", "T5H", "T5V"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(18).Trim
                                                Case "R"
                                                    '最小チェック値設定
                                                    selListMin.Clear()  '※上記で設定されている場合、最小チェック値を塗替える
                                                    selListMin.Add("5:*")
                                                Case "D"
                                                    '最小チェック値設定
                                                    selListMin.Clear()  '※上記で設定されている場合、最小チェック値を塗替える
                                                    selListMin.Add("10:*")
                                            End Select
                                    End Select
                            End Select
                    End Select
                    '↑RM1212080 2012/12/05 Y.Tachi
                Case "L4"

                    'スイッチ
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                        Case ""
                            'Do Nothing
                        Case "T0H", "T0V", "T5H", "T5V"
                            '最小チェック値設定
                            selListMin.Add("10:12,16")
                            selListMin.Add("5:20,25,32,40,50,63,80,100")
                        Case "F2V", "F3V", "F2YV", "F3YV"
                            '最小チェック値設定
                            selListMin.Add("10:20")
                            selListMin.Add("5:12,16,25,32,40,50,63,80,100")
                        Case "T2H", "T2V", "T3H", "T3V", "T3PH", "T3PV", "F2H", "F2V", "F2YH", "F3YH"
                            '最小チェック値設定
                            selListMin.Add("5:*")
                        Case Else
                            '最小チェック値設定
                            selListMin.Add("10:*")
                    End Select

                    'バリエーション②
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "K"
                            '最小チェック値設定
                            selListMin.Clear()  '※上記で設定されている場合、最小チェック値を塗替える
                            selListMin.Add("20:*")

                        Case "KG1"
                            '最小チェック値設定
                            selListMin.Clear()  '※上記で設定されている場合、最小チェック値を塗替える
                            selListMin.Add("20:*")

                    End Select
            End Select

            '最小ストロークチェック
            If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                          objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                          selListMin, 2) = False Then
                intKtbnStrcSeqNo = 14
                strMessageCd = "W0200"
                fncHighLoadBaseCheck = False
                Exit Try
            End If

            '2010/11/02 ADD RM1011020(12月VerUP:SSD2シリーズ) <---END

            'スイッチ毎のチェック
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 5
                strMessageCd = "W0200"
                fncHighLoadBaseCheck = False
                Exit Try
            End If

            '2010/12/10 ADD RM1012055(1月VerUP:SSD2シリーズ) START--->
            '*-----<< オプション「中間ストローク専用本体」チェック >>-----*
            Dim strOp() As String
            strOp = Split(objKtbnStrc.strcSelection.strOpSymbol(19).Trim, ",")
            If Not fncOptionSCheck(strOp, _
                                objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                objKtbnStrc.strcSelection.strOpSymbol(14).Trim) Then
                intKtbnStrcSeqNo = 14
                strMessageCd = "W0830"
                fncHighLoadBaseCheck = False
                Exit Try
            End If
            '2010/12/10 ADD RM1012055(1月VerUP:SSD2シリーズ) <---END

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncHighLoadBaseCheck
    '*【処理】
    '*  高荷重ベースチェック
    '*【概要】
    '*  高荷重ベースをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncHighLoadBaseP4Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Try

            fncHighLoadBaseP4Check = True

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            'スイッチ毎のチェック
            'Ｓ１
            'RM1305005 2013/05/30 ローカル版と差異修正
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(16).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(18).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 5
                strMessageCd = "W0200"
                fncHighLoadBaseP4Check = False
                Exit Try
            End If

            '2010/11/02 DEL RM1011020(12月VerUP:SSD2シリーズ) START--->
            ''RM0906034 2009/08/05 Y.Miura　二次電池対応機種追加
            'If objKtbnStrc.strcSelection.strKeyKataban.Equals("L") Then
            '2010/11/02 DEL RM1011020(12月VerUP:SSD2シリーズ) <---END

            '基本ベースチェック　二次電池対応
            '↓2012/01/05 RM1201XXX intOptionPos(9→19)変更
            If fncP4Check(objKtbnStrc, _
                                    intKtbnStrcSeqNo, _
                                    strOptionSymbol, _
                                    strMessageCd, _
                                    19) = False Then
                fncHighLoadBaseP4Check = False
                Exit Try
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncLongBaseP4Check
    '*【処理】
    '*  ロングストローク（両ロッド）（Ｐ４）ベースチェック
    '*【概要】
    '*  ロングストローク（両ロッド）（Ｐ４）をチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncLongBaseP4Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Try

            fncLongBaseP4Check = True

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            'スイッチ毎のチェック
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 5
                strMessageCd = "W0200"
                fncLongBaseP4Check = False
                Exit Try
            End If

            '基本ベースチェック　二次電池対応
            If fncP4Check(objKtbnStrc, _
                                    intKtbnStrcSeqNo, _
                                    strOptionSymbol, _
                                    strMessageCd, _
                                    9) = False Then
                fncLongBaseP4Check = False
                Exit Try
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '2010/11/02 DEL RM1011020(12月VerUP:SSD2シリーズ) START--->
    ''********************************************************************************************
    ''*【関数名】
    ''*  fncNonRotatingBaseCheck
    ''*【処理】
    ''*  回り止めベースチェック
    ''*【概要】
    ''*  回り止めベースをチェックする
    ''*【引数】
    ''*  <Object>       objKtbnStrc          引当形番情報
    ''*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    ''*  <String>       strOptionSymbol      オプション記号
    ''*  <String>       strMessageCd         メッセージコード
    ''*【戻り値】
    ''*  <Boolean>
    ''********************************************************************************************
    'Private Function fncNonRotatingBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
    '                                      ByRef intKtbnStrcSeqNo As Integer, _
    '                                      ByRef strOptionSymbol As String, _
    '                                      ByRef strMessageCd As String) As Boolean

    '    Try

    '        fncNonRotatingBaseCheck = True

    '        '*-----<< Ⅰ．最小ストロークチェック >>-----*
    '        'スイッチ毎のチェック
    '        If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = False Then
    '            intKtbnStrcSeqNo = 5
    '            strMessageCd = "W0200"
    '            fncNonRotatingBaseCheck = False
    '            Exit Try
    '        End If

    '    Catch ex As Exception

    '        Throw ex

    '    End Try

    'End Function

    ''********************************************************************************************
    ''*【関数名】
    ''*  fncPositionLockingBaseCheck
    ''*【処理】
    ''*  落下防止ベースチェック
    ''*【概要】
    ''*  落下防止ベースをチェックする
    ''*【引数】
    ''*  <Object>       objKtbnStrc          引当形番情報
    ''*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    ''*  <String>       strOptionSymbol      オプション記号
    ''*  <String>       strMessageCd         メッセージコード
    ''*【戻り値】
    ''*  <Boolean>
    ''********************************************************************************************
    'Private Function fncPositionLockingBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
    '                                      ByRef intKtbnStrcSeqNo As Integer, _
    '                                      ByRef strOptionSymbol As String, _
    '                                      ByRef strMessageCd As String) As Boolean

    '    Try

    '        fncPositionLockingBaseCheck = True

    '        '*-----<< Ⅰ．最小ストロークチェック >>-----*
    '        'スイッチ毎のチェック
    '        If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = False Then
    '            intKtbnStrcSeqNo = 5
    '            strMessageCd = "W0200"
    '            fncPositionLockingBaseCheck = False
    '            Exit Try
    '        End If

    '    Catch ex As Exception

    '        Throw ex

    '    End Try

    'End Function

    ''********************************************************************************************
    ''*【関数名】
    ''*  fncDoubleRodBaseCheck
    ''*【処理】
    ''*  押出しベースチェック
    ''*【概要】
    ''*  押出しベースをチェックする
    ''*【引数】
    ''*  <Object>       objKtbnStrc          引当形番情報
    ''*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    ''*  <String>       strOptionSymbol      オプション記号
    ''*  <String>       strMessageCd         メッセージコード
    ''*【戻り値】
    ''*  <Boolean>
    ''********************************************************************************************
    'Private Function fncSpringReturnBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
    '                                      ByRef intKtbnStrcSeqNo As Integer, _
    '                                      ByRef strOptionSymbol As String, _
    '                                      ByRef strMessageCd As String) As Boolean

    '    Try

    '        fncSpringReturnBaseCheck = True

    '        '*-----<< Ⅰ．最小ストロークチェック >>-----*
    '        'スイッチ毎のチェック
    '        If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = False Then
    '            intKtbnStrcSeqNo = 5
    '            strMessageCd = "W0200"
    '            fncSpringReturnBaseCheck = False
    '            Exit Try
    '        End If

    '    Catch ex As Exception

    '        Throw ex

    '    End Try

    'End Function

    ''********************************************************************************************
    ''*【関数名】
    ''*  fncSpringExtendBaseCheck
    ''*【処理】
    ''*  引込みベースチェック
    ''*【概要】
    ''*  引込みベースをチェックする
    ''*【引数】
    ''*  <Object>       objKtbnStrc          引当形番情報
    ''*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    ''*  <String>       strOptionSymbol      オプション記号
    ''*  <String>       strMessageCd         メッセージコード
    ''*【戻り値】
    ''*  <Boolean>
    ''********************************************************************************************
    'Private Function fncSpringExtendBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
    '                                      ByRef intKtbnStrcSeqNo As Integer, _
    '                                      ByRef strOptionSymbol As String, _
    '                                      ByRef strMessageCd As String) As Boolean

    '    Try

    '        fncSpringExtendBaseCheck = True

    '        '*-----<< Ⅰ．最小ストロークチェック >>-----*
    '        'スイッチ毎のチェック
    '        If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = False Then
    '            intKtbnStrcSeqNo = 5
    '            strMessageCd = "W0200"
    '            fncSpringExtendBaseCheck = False
    '            Exit Try
    '        End If

    '    Catch ex As Exception

    '        Throw ex

    '    End Try

    'End Function
    '2010/11/02 DEL RM1011020(12月VerUP:SSD2シリーズ) <---END

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
    '*  fncSSD2SwitchStrokeCheck
    '*【処理】
    '*  スイッチ毎のチェック
    '*【概要】
    '*  スイッチ形番毎にストロークをチェックする
    '*【引数】
    '*  <String>        strStroke           ストローク
    '*  <String>        strSwitchKataban    スイッチ形番
    '*  <String>        strSwitchQty        スイッチ数
    '*  <String>        strVariation        バリエーション
    '*  <String>        strPortSize         口径
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncSSD2SwitchStrokeCheck(ByVal strKeyKataban As String, _
                                              ByVal strStroke As String, _
                                              ByVal strSwitchKataban As String, _
                                              ByVal strSwitchQty As String, _
                                              ByVal strVariation As String, _
                                              ByVal strPortSize As String, _
                                              ByVal strSwitch As String)

        Dim objPrice As New KHUnitPrice

        Try

            fncSSD2SwitchStrokeCheck = False

            '↓RM1212080 2012/12/04 Y.Tachi 
            If strKeyKataban = "4" Then
                If InStr(1, strSwitch, "L") = 0 Then
                Else
                    If strSwitchKataban.Length = 0 Then
                    Else
                        Select Case strSwitchKataban.Trim
                            Case "SW17", "SW20", "SW27", "SW28", "SW29", "SW30", "SW69", "SW70", "SWAK"
                                Select Case strPortSize.Trim
                                    Case "12", "16"
                                        Select Case strVariation
                                            Case "", "X", "Y", "O", "B", "W", "M"
                                                If strSwitchQty.Trim = "R" Then
                                                    If Val(strStroke) < 5 Then
                                                        Exit Try
                                                    End If
                                                Else
                                                    If Val(strStroke) < 10 Then
                                                        Exit Try
                                                    End If
                                                End If
                                        End Select
                                End Select
                        End Select
                    End If
                End If
            End If
            If strKeyKataban = "L" Then
                If InStr(1, strSwitch, "L") = 0 Then
                Else
                    If strSwitchKataban.Length = 0 Then
                    Else
                        Select Case strSwitchKataban.Trim
                            Case "SW17", "SW20", "SW27", "SW28", "SW29", "SW30", "SW69", "SW70", "SWAK"
                                Select Case strPortSize.Trim
                                    Case "12", "16"
                                        If strVariation.Trim = "K" And _
                                           strSwitchQty.Trim = "R" Then
                                            If Val(strStroke) < 5 Then
                                                Exit Try
                                            End If
                                        Else
                                            If Val(strStroke) < 10 Then
                                                Exit Try
                                            End If
                                        End If
                                End Select
                        End Select
                    End If
                End If
            End If
            '↑RM1212080 2012/12/04 Y.Tachi

            'SW選択有無判定
            If strSwitchKataban.Trim = "" Then
            Else
                Select Case strSwitchKataban.Trim
                    'RM0906034 2009/08/05 Y.Miura　二次電池対応
                    'Case "T0H", "T0V", "T5H", "T5V"
                    Case "T0H", "T0V", "T5H", "T5V", "SW27"
                        Select Case strPortSize.Trim
                            Case "12", "16"
                                Select Case strKeyKataban
                                    'RM0906034 2009/08/05 Y.Miura　二次電池対応
                                    'Case "", "K", "X", "Y"
                                    Case "", "K", "X", "Y", "4"
                                        Select Case strVariation.Trim
                                            Case "", "X", "Y", "O", "B", "W", "M"
                                                If strSwitchQty.Trim = "R" Then
                                                    If Val(strStroke) < 5 Then
                                                        Exit Try
                                                    End If
                                                Else
                                                    If Val(strStroke) < 10 Then
                                                        Exit Try
                                                    End If
                                                End If
                                            Case Else
                                                If CInt(strStroke) < 10 Then
                                                    Exit Try
                                                End If
                                        End Select
                                        '↓RM1212080 2012/12/04 Y.Tachi
                                    Case "K"
                                        If strVariation.Trim = "K" And _
                                           strSwitchQty.Trim = "R" Then
                                            If Val(strStroke) < 5 Then
                                                Exit Try
                                            End If
                                        Else
                                            If Val(strStroke) < 5 Then
                                                Exit Try
                                            End If
                                        End If
                                    Case "D"
                                        If strVariation.Trim = "DM" Then
                                            If strSwitchQty.Trim = "R" Then
                                                If Val(strStroke) < 5 Then
                                                    Exit Try
                                                End If
                                            Else
                                                If Val(strStroke) < 10 Then
                                                    Exit Try
                                                End If
                                            End If
                                        End If
                                        '↑RM1212080 2012/12/04 Y.Tachi
                                End Select
                            Case Else
                                If CInt(strStroke) < 5 Then
                                    Exit Try
                                End If
                        End Select
                        'RM0906034 2009/08/05 Y.Miura　二次電池対応
                        'Case "F2V", "F3V", "F2YV", "F3YV"
                    Case "F2V", "F3V", "F2YV", "F3YV", "SW83", "SW84", "SW87", "SW88"
                        Select Case strPortSize.Trim
                            Case "20"
                                Select Case strKeyKataban
                                    Case "", "D", "M", "X", "Y", "E"
                                        If CInt(strStroke) < 15 Then
                                            Exit Try
                                        End If
                                        'RM0906034 2009/08/05 Y.Miura　二次電池対応
                                        'Case "K"
                                    Case "K", "L"
                                        If CInt(strStroke) < 10 Then
                                            Exit Try
                                        End If
                                End Select
                            Case Else
                                If CInt(strStroke) < 5 Then
                                    Exit Try
                                End If
                        End Select
                        'RM0906034 2009/08/05 Y.Miura　二次電池対応
                        'Case "T2H", "T2V", "T3H", "T3V", "T3PH", "T3PV", "F2H", "F3H", "F2YH", "F3YH"
                        'RM1210067 2013/02/01 Y.Tachi ローカル版との差異修正
                        'RM1305005 2013/05/30 ローカル版との差異修正
                    Case "T2H", "T2V", "T3H", "T3V", "T3PH", "T3PV", "F2H", "F3H", "F2YH", "F3YH", "F3PH", "F3PV", _
                        "SW11", "SW12", "SW13", "SW14", "SW15", "SW16", "SW21", "SW22", "SW23", "SW24", "SW25", "SW26", "SW27", _
                        "SW81", "SW82", "SW83", "SW84", "SW85", "SW86", "SW87", "SW88"
                        If CInt(strStroke) < 5 Then
                            Exit Try
                        End If
                    Case Else
                        If CInt(strStroke) < 10 Then
                            Exit Try
                        End If
                End Select
            End If

            fncSSD2SwitchStrokeCheck = True

        Catch ex As Exception

            Throw ex

        Finally

            objPrice = Nothing

        End Try

    End Function

    '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) START--->
    '********************************************************************************************
    '*【関数名】
    '*  fncSSD2BaseStrokeCheck
    '*【処理】
    '*  スイッチ毎のチェック(基本ベース用)
    '*【概要】
    '*  スイッチ形番毎にストロークをチェックする
    '*【引数】
    '*  <String>        strStroke           ストローク
    '*  <String>        strPortSize         口径
    '*  <String>        lstCheck　　　　　　チェックリスト
    '*  <Integer>       checkFlg　　　　　　チェックフラグ(1:より大きい、2:より小さい)
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncSSD2BaseStrokeCheck(ByVal strStroke As String, _
                                            ByVal strPortSize As String, _
                                            ByVal lstCheck As ArrayList, _
                                            ByVal checkFlg As Integer)

        Dim wkSp() As String

        Try

            fncSSD2BaseStrokeCheck = False

            If strStroke.Equals(String.Empty) Then
                Exit Try
            End If

            For i As Integer = 0 To lstCheck.Count - 1
                '分割
                wkSp = Split(lstCheck.Item(i).ToString, ":")

                '口径
                '2010/11/02 MOD RM1011020(12月VerUP:SSD2シリーズ) START--->
                If InStr(wkSp(1), strPortSize) > 0 Or wkSp(1) = "*" Then
                    'If InStr(wkSp(1), strPortSize) > 0 Then
                    '2010/11/02 MOD RM1011020(12月VerUP:SSD2シリーズ) <---END
                    'ストローク
                    Select Case checkFlg
                        Case 1
                            If strStroke > CInt(wkSp(0)) Then
                                Exit Try
                            End If
                        Case 2
                            If strStroke < CInt(wkSp(0)) Then
                                Exit Try
                            End If

                    End Select
                End If
            Next

            fncSSD2BaseStrokeCheck = True

        Catch ex As Exception
            Throw ex

        End Try
    End Function
    '2010/10/05 ADD RM1010017(11月VerUP:SSD2シリーズ) <---END
    '2010/12/10 ADD RM1012055(1月VerUP:SSD2シリーズ) START--->
    ''' <summary>
    ''' オプション「中間ストローク専用本体」チェック
    ''' </summary>
    ''' <param name="strOp">オプションリスト</param>
    ''' <param name="strBoreSize">口径</param>
    ''' <param name="strStroke">S2ストローク</param>
    ''' <returns>True:成功、False:失敗</returns>
    ''' <remarks></remarks>
    Private Function fncOptionSCheck(ByVal strOp() As String, _
                                    ByVal strBoreSize As String, ByVal strStroke As String) As Boolean
        Try
            Dim ret As Boolean = True
            For i As Integer = 0 To strOp.Length - 1
                Select Case Trim(strOp(i))
                    Case "S"
                        Select Case strBoreSize
                            Case "12", "16"
                                Select Case strStroke
                                    Case "5", "10", "15", "20", "25", "30"
                                        ret = False
                                        Exit For
                                End Select
                            Case "20", "25"
                                Select Case strStroke
                                    Case "5", "10", "15", "20", "25", "30", "35", "40", "45", "50"
                                        ret = False
                                        Exit For
                                End Select
                            Case "32", "40"
                                Select Case strStroke
                                    Case "5", "10", "15", "20", "25", "30", "35", "40", "45", "50", "75", "100"
                                        ret = False
                                        Exit For
                                End Select
                            Case "50", "63", "80", "100"
                                Select Case strStroke
                                    Case "10", "15", "20", "25", "30", "35", "40", "45", "50", "75", "100"
                                        ret = False
                                        Exit For
                                End Select
                        End Select
                End Select
            Next

            Return ret

        Catch ex As Exception
            Throw ex

        End Try

    End Function
    '2010/12/10 ADD RM1012055(1月VerUP:SSD2シリーズ) <---END
End Module
