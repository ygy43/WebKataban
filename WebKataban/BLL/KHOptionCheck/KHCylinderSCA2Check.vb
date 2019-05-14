Module KHCylinderSCA2Check

#Region " Definition "

    Private bolC5CheckFlg As Boolean                'C5判定フラグ

#End Region

    '********************************************************************************************
    '*【関数名】
    '*  fncCheckSelectOption
    '*【処理】
    '*  シリンダチェック
    '*【概要】
    '*  シリンダＳＣＡ２シリーズをチェックする
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

            'C5チェック
            bolC5CheckFlg = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc)

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "SCA2"
                    '基本ベース毎にチェック
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        Case "", "2"
                            '基本ベースチェック
                            If fncStandardBaseCheck(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case "V"
                            'バルブ付ベースチェック
                            If fncValveBaseCheck(objKtbnStrc, _
                                                 intKtbnStrcSeqNo, _
                                                 strOptionSymbol, _
                                                 strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case "B", "C"
                            '背合わせ＆２段形ベースチェック
                            If fncDoubleRodBaseCheck(objKtbnStrc, _
                                                     intKtbnStrcSeqNo, _
                                                     strOptionSymbol, _
                                                     strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case "D", "E"
                            '両ロッドベースチェック
                            If fncHighLoadBaseCheck(objKtbnStrc, _
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
    '*【更新履歴】
    '*                                          更新日：2008/01/10      更新者：NII A.Takahashi
    '*  ・最小ストロークの変更に伴い、バリエーション毎に最小ストロークチェックをするように修正
    '********************************************************************************************
    Private Function fncStandardBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False
        Dim bolOptionB1 As Boolean = False
        Dim bolOptionB2 As Boolean = False
        Dim bolOptionB3 As Boolean = False
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncStandardBaseCheck = True

            'オプション選択チェック
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "2"
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(15), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "I"
                                bolOptionI = True
                            Case "Y"
                                bolOptionY = True
                            Case "B1"
                                bolOptionB1 = True
                            Case "B2"
                                bolOptionB2 = True
                            Case "B3"
                                bolOptionB3 = True
                        End Select
                    Next
                Case Else
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(14), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "I"
                                bolOptionI = True
                            Case "Y"
                                bolOptionY = True
                            Case "B1"
                                bolOptionB1 = True
                            Case "B2"
                                bolOptionB2 = True
                            Case "B3"
                                bolOptionB3 = True
                        End Select
                    Next
            End Select
          

            'バリエーション「Q2」＋ジャバラ「J」「L」の組合せは原価積算対応
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q2") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("J") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("L") >= 0 Then
                    intKtbnStrcSeqNo = 13
                    'RM1210067 2013/02/01 Y.Tachi ローカル版との差異修正(W0710→W0720)
                    strMessageCd = "W0720"
                    fncStandardBaseCheck = False
                    Exit Try
                End If
            End If

            'ジャバラ「Ｊ」「Ｌ」の最大ストロークは2000
            If objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("J") >= 0 Or _
               objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("L") >= 0 Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    Case "63", "80", "100"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 2000 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                End Select
            End If

            '「B1」(一山ブラケット)組合せチェック
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "2"
                    If bolOptionB1 = True Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CB" Or bolOptionY = True Then
                        Else
                            intKtbnStrcSeqNo = 15
                            strMessageCd = "W0290"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If

                    '「B2」(ニ山ブラケット)組合せチェック
                    If bolOptionB2 = True Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CA" Or bolOptionI = True Then
                        Else
                            intKtbnStrcSeqNo = 15
                            strMessageCd = "W0300"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If

                    '「B3」(一山ブラケット)組合せチェック
                    If bolOptionB3 = True Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CB" Or bolOptionY = True Then
                        Else
                            intKtbnStrcSeqNo = 15
                            strMessageCd = "W0310"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If
                Case Else
                    If bolOptionB1 = True Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CB" Or bolOptionY = True Then
                        Else
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0290"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If

                    '「B2」(ニ山ブラケット)組合せチェック
                    If bolOptionB2 = True Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CA" Or bolOptionI = True Then
                        Else
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0300"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If

                    '「B3」(一山ブラケット)組合せチェック
                    If bolOptionB3 = True Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CB" Or bolOptionY = True Then
                        Else
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0310"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If
            End Select



            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "P", "R", "PH", "RO", "RU", "RG", "RG1", "RG2", "RG3", "RG4"
                    If fncVarPRMinStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(10).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(11).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(12).Trim) = False Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                Case "Q2"
                    If fncVarQMinStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(10).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(11).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(12).Trim) = False Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                Case "T", "T1", "PK", "PKH", "RK", "RKO", "RKG", "RKG1", "RKG4", _
                     "Q2K", "KH", "KT", "KT1", "KT2", "KO", "KG", "KG1", "KG4", "KTG1", _
                     "KT1G1", "KT2G1", "TG1", "T1G1"
                    If fncVarKMinStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(10).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(11).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(12).Trim) = False Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                Case Else
                    '基本チェック
                    If fncStdMinStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(10).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(11).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(12).Trim) = False Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
            End Select

            '*-----<< Ⅱ．最大ストロークチェック >>-----*
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "2"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "", "G", "G2", "G3"
                            '基本STチェック
                            If fncStdMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                            objKtbnStrc.strcSelection.strOpSymbol(7).Trim) = False Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If

                        Case "B", "W"
                            If fncVarPRMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                      objKtbnStrc.strcSelection.strOpSymbol(7).Trim) = False Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        Case "P", "R"
                            If fncVarPRMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                      objKtbnStrc.strcSelection.strOpSymbol(7).Trim) = False Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        Case "K"
                            If fncVarKMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                     objKtbnStrc.strcSelection.strOpSymbol(7).Trim) = False Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                    End Select

                Case Else
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "", "R", "Q2", "H", "T", "T1", "T2", "G", "G1", "G2", "G3", "G4", _
                             "RG", "RG1", "RG2", "RG3", "RG4", "TG1", "T1G1", "T2G1", "T2G4"
                            '基本STチェック
                            If fncStdMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                            objKtbnStrc.strcSelection.strOpSymbol(7).Trim) = False Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If

                        Case "O", "U", "RO", "RU", "RKO", "KO"
                            If fncVarOUMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                      objKtbnStrc.strcSelection.strOpSymbol(7).Trim) = False Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        Case "P", "PK", "PH", "PKH"
                            If fncVarDPMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                      objKtbnStrc.strcSelection.strOpSymbol(7).Trim) = False Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        Case "K", "RK", "RKG", "RKG1", "RKG4", "Q2K", "KH", "KT", "KT1", _
                             "KG", "KG1", "KG4", "KTG1", "KT1G1", "KT2G1"
                            If fncVarKMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                     objKtbnStrc.strcSelection.strOpSymbol(7).Trim) = False Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                    End Select
            End Select

            '2012/07/27 オプション外チェック
            If objKtbnStrc.strcSelection.strOtherOption.Trim <> "" Then
                '支持金具90°回転(K1)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K1") >= 0 Then
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "00", "FC"
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0430"
                            fncStandardBaseCheck = False
                            Exit Try
                    End Select
                End If

                '支持金具180°回転(K2)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K2") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "LB" Then
                        intKtbnStrcSeqNo = 3
                        strMessageCd = "W0440"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                End If

                '支持金具270°回転(K3)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K3") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "LB" Then
                        intKtbnStrcSeqNo = 3
                        strMessageCd = "W0450"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                End If

                'P5
                Select Case objKtbnStrc.strcSelection.strKeyKataban
                    Case "2"
                        If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P5") >= 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "CB" And _
                               objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("Y") < 0 Then
                                intKtbnStrcSeqNo = 15
                                strMessageCd = "W0470"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        End If
                    Case Else
                        If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P5") >= 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "CB" And _
                               objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("Y") < 0 Then
                                intKtbnStrcSeqNo = 14
                                strMessageCd = "W0470"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        End If
                End Select

                'M1
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("M1") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("K") >= 0 Then
                        intKtbnStrcSeqNo = 1
                        strMessageCd = "W0760"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                End If

                'J9
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("J9") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("J") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("K") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("O") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("U") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T") >= 0 Then
                        intKtbnStrcSeqNo = 13
                        strMessageCd = "W0770"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                End If

                'T9
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("T9") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("O") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("U") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("H") >= 0 Then
                        intKtbnStrcSeqNo = 13
                        strMessageCd = "W0770"
                        fncStandardBaseCheck = False
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
    '*  fncValveBaseCheck
    '*【処理】
    '*  バルブ付ベースチェック
    '*【概要】
    '*  バルブ付ベースをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2008/01/10      更新者：NII A.Takahashi
    '*  ・最小ストロークの変更に伴い、バリエーション毎に最小ストロークチェックをするように修正
    '********************************************************************************************
    Private Function fncValveBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                       ByRef intKtbnStrcSeqNo As Integer, _
                                       ByRef strOptionSymbol As String, _
                                       ByRef strMessageCd As String) As Boolean

        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False
        Dim bolOptionB1 As Boolean = False
        Dim bolOptionB2 As Boolean = False
        Dim bolOptionB3 As Boolean = False
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncValveBaseCheck = True

            'オプション選択チェック
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(14), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "I"
                        bolOptionI = True
                    Case "Y"
                        bolOptionY = True
                    Case "B1"
                        bolOptionB1 = True
                    Case "B2"
                        bolOptionB2 = True
                    Case "B3"
                        bolOptionB3 = True
                End Select
            Next

            'ジャバラ「Ｊ」「Ｌ」の最大ストロークは2000
            If objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("J") >= 0 Or _
               objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("L") >= 0 Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    Case "63", "80", "100"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 2000 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncValveBaseCheck = False
                            Exit Try
                        End If
                End Select
            End If

            'オプションJ,L最大ストロークチェック
            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(13).Trim, "J") <> 0 Or _
               InStr(1, objKtbnStrc.strcSelection.strOpSymbol(13).Trim, "L") <> 0 Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    Case "63", "80", "100"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 2000 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncValveBaseCheck = False
                            Exit Try
                        End If
                End Select
            End If

            '「B1」(一山ブラケット)組合せチェック
            If bolOptionB1 = True Then
                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CB" Or bolOptionY = True Then
                Else
                    intKtbnStrcSeqNo = 14
                    strMessageCd = "W0290"
                    fncValveBaseCheck = False
                    Exit Try
                End If
            End If

            '「B2」(ニ山ブラケット)組合せチェック
            If bolOptionB2 = True Then
                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CA" Or bolOptionI = True Then
                Else
                    intKtbnStrcSeqNo = 14
                    strMessageCd = "W0300"
                    fncValveBaseCheck = False
                    Exit Try
                End If
            End If

            '「B3」(一山ブラケット)組合せチェック
            If bolOptionB3 = True Then
                If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CB" Or bolOptionY = True Then
                Else
                    intKtbnStrcSeqNo = 14
                    strMessageCd = "W0310"
                    fncValveBaseCheck = False
                    Exit Try
                End If
            End If

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "V1", "V2", "V", "PV1", "PV2", "PV", "RV1", "RV2", "RV", "RVK", "RV1G", "RV2G", "RVG", _
                     "RV1G1", "RV2G1", "RVG1", "RV1G4", "RV2G4", "RVG4", "V1G", "V2G", "VG", "V1G1", _
                     "V2G1", "VG1", "V1G4", "V2G4", "VG4"
                    If fncVarVMinStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(10).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(11).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(12).Trim) = False Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0200"
                        fncValveBaseCheck = False
                        Exit Try
                    End If
                Case "PV1K", "PV2K", "PVK", "RV1K", "RV2K", "RV1KG", "RV2KG", "RVKG", "RV1KG1", _
                     "RV2KG1", "RVKG1", "RV1KG4", "RV2KG4", "RVKG4", "V1K", "V2K", "VK", "V1KG", _
                     "V2KG", "VKG", "V1KG1", "V2KG1", "VKG1", "V1KG4", "V2KG4", "VKG4"
                    If fncVarKMinStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(10).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(11).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(12).Trim) = False Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0200"
                        fncValveBaseCheck = False
                        Exit Try
                    End If
            End Select

            '*-----<< Ⅱ．最大ストロークチェック >>-----*
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "V1", "V2", "V", "RV1", "RV2", "RV", "RV1K", "RV2K", "RVK", "RV1G", "RV2G", _
                     "RVG", "RV1G1", "RV2G1", "RVG1", "RV1G4", "RV2G4", "RVG4", "V1G", "V2G", "VG", _
                     "V1G1", "V2G1", "VG1", "V1G4", "V2G4", "VG4"
                    If fncStdMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(7).Trim) = False Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0200"
                        fncValveBaseCheck = False
                        Exit Try
                    End If
                Case "PV1", "PV2", "PV", "PV1K", "PV2K", "PVK"
                    If fncVarDPMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(7).Trim) = False Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0200"
                        fncValveBaseCheck = False
                        Exit Try
                    End If
                Case "RV1KG", "RV2KG", "RVKG", "RV1KG1", "RV2KG1", "RVKG1", "RV1KG4", "RV2KG4", "RVKG4", _
                     "V1K", "V2K", "VK", "V1KG", "V2KG", "VKG", "V1KG1", "V2KG1", "VKG1", "V1KG4", "V2KG4", "VKG4"
                    If fncVarKMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(7).Trim) = False Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0200"
                        fncValveBaseCheck = False
                        Exit Try
                    End If
            End Select

            '2012/07/27 オプション外チェック
            If objKtbnStrc.strcSelection.strOtherOption.Trim <> "" Then
                '支持金具180°回転(K2)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K2") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "LB" Then
                        intKtbnStrcSeqNo = 3
                        strMessageCd = "W0440"
                        fncValveBaseCheck = False
                        Exit Try
                    End If
                End If

                '支持金具270°回転(K3)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K3") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "LB" Then
                        intKtbnStrcSeqNo = 3
                        strMessageCd = "W0450"
                        fncValveBaseCheck = False
                        Exit Try
                    End If
                End If

                'トラニオン位置
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("AQ") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "TC" And _
                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "TF" Then
                        intKtbnStrcSeqNo = 3
                        strMessageCd = "W0460"
                        fncValveBaseCheck = False
                        Exit Try
                    End If
                End If

                'P5
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P5") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "CB" And _
                       objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("Y") < 0 Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0470"
                        fncValveBaseCheck = False
                        Exit Try
                    End If
                End If

                'M1
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("M1") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("K") >= 0 Then
                        intKtbnStrcSeqNo = 1
                        strMessageCd = "W0760"
                        fncValveBaseCheck = False
                        Exit Try
                    End If
                End If

                'J9
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("J9") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("J") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("K") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("L") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G") >= 0 Then
                        intKtbnStrcSeqNo = 13
                        strMessageCd = "W0770"
                        fncValveBaseCheck = False
                        Exit Try
                    End If
                End If

                'T9
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("T9") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G") >= 0 Then
                        intKtbnStrcSeqNo = 1
                        strMessageCd = "W0770"
                        fncValveBaseCheck = False
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
    '*  fncDoubleRodBaseCheck
    '*【処理】
    '*  背合わせ＆２段形ベースチェック
    '*【概要】
    '*  背合わせ＆２段形ベースをチェックする
    '*【引数】
    '*  <Object>       objKtbnStrc          引当形番情報
    '*  <Integer>      intKtbnStrcSeqNo     形番構成順序
    '*  <String>       strOptionSymbol      オプション記号
    '*  <String>       strMessageCd         メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2008/01/10      更新者：NII A.Takahashi
    '*  ・最小ストロークの変更に伴い、バリエーション毎に最小ストロークチェックをするように修正
    '********************************************************************************************
    Private Function fncDoubleRodBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                           ByRef intKtbnStrcSeqNo As Integer, _
                                           ByRef strOptionSymbol As String, _
                                           ByRef strMessageCd As String) As Boolean

        Dim bolOptionI As Boolean = False
        Dim bolOptionY As Boolean = False
        Dim bolOptionB1 As Boolean = False
        Dim bolOptionB2 As Boolean = False
        Dim bolOptionB3 As Boolean = False
        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            fncDoubleRodBaseCheck = True

            'オプション選択チェック
            'strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(14), CdCst.Sign.Delimiter.Comma)  'RM1003086 不具合修正
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "C"
                    '食品製造工程向け商品
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(19), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "I"
                                bolOptionI = True
                            Case "Y"
                                bolOptionY = True
                            Case "B1"
                                bolOptionB1 = True
                            Case "B2"
                                bolOptionB2 = True
                            Case "B3"
                                bolOptionB3 = True
                        End Select
                    Next
                Case Else
                    strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(18), CdCst.Sign.Delimiter.Comma)
                    For intLoopCnt = 0 To strOpArray.Length - 1
                        Select Case strOpArray(intLoopCnt).Trim
                            Case "I"
                                bolOptionI = True
                            Case "Y"
                                bolOptionY = True
                            Case "B1"
                                bolOptionB1 = True
                            Case "B2"
                                bolOptionB2 = True
                            Case "B3"
                                bolOptionB3 = True
                        End Select
                    Next

            End Select

            'ジャバラ「Ｊ」「Ｌ」の最大ストロークは2000
            If objKtbnStrc.strcSelection.strOpSymbol(17).IndexOf("J") >= 0 Or _
               objKtbnStrc.strcSelection.strOpSymbol(17).IndexOf("L") >= 0 Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    Case "63", "80", "100"
                        'S1
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 2000 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncDoubleRodBaseCheck = False
                            Exit Try
                        End If
                        'S2
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(13).Trim) > 2000 Then
                            intKtbnStrcSeqNo = 13
                            strMessageCd = "W0200"
                            fncDoubleRodBaseCheck = False
                            Exit Try
                        End If
                End Select
            End If


            '「B1」(一山ブラケット)組合せチェック
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "C"
                    If bolOptionB1 = True Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CB" Or bolOptionY = True Then
                        Else
                            intKtbnStrcSeqNo = 19
                            strMessageCd = "W0290"
                            fncDoubleRodBaseCheck = False
                            Exit Try
                        End If
                    End If

                    '「B2」(一山ブラケット)組合せチェック
                    If bolOptionB2 = True Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CA" Or bolOptionI = True Then
                        Else
                            intKtbnStrcSeqNo = 19
                            strMessageCd = "W0300"
                            fncDoubleRodBaseCheck = False
                            Exit Try
                        End If
                    End If

                    '「B3」(一山ブラケット)組合せチェック
                    If bolOptionB3 = True Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CB" Or bolOptionY = True Then
                        Else
                            intKtbnStrcSeqNo = 19
                            strMessageCd = "W0310"
                            fncDoubleRodBaseCheck =
                                False
                            Exit Try
                        End If
                    End If
                Case Else
                    If bolOptionB1 = True Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CB" Or bolOptionY = True Then
                        Else
                            intKtbnStrcSeqNo = 18
                            strMessageCd = "W0290"
                            fncDoubleRodBaseCheck = False
                            Exit Try
                        End If
                    End If

                    '「B2」(一山ブラケット)組合せチェック
                    If bolOptionB2 = True Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CA" Or bolOptionI = True Then
                        Else
                            intKtbnStrcSeqNo = 18
                            strMessageCd = "W0300"
                            fncDoubleRodBaseCheck = False
                            Exit Try
                        End If
                    End If

                    '「B3」(一山ブラケット)組合せチェック
                    If bolOptionB3 = True Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim = "CB" Or bolOptionY = True Then
                        Else
                            intKtbnStrcSeqNo = 18
                            strMessageCd = "W0310"
                            fncDoubleRodBaseCheck = False
                            Exit Try
                        End If
                    End If

            End Select

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "BK", "BT", "BT1", "BKH", "BKT", "BKT1", "BKT2", "BKO", "BKG", "BKG1", _
                     "BKG4", "BKTG1", "BKT1G1", "BKT2G1", "BTG1", "BT1G1", _
                     "WK", "WT", "WT1", "WKH", "WKT", "WKT1", "WKT2", "WKG", "WKG1", _
                     "WKG4", "WKTG1", "WKT1G1", "WTG1", "WT1G1"
                    'S1
                    If fncVarKMinStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(10).Trim) = False Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0200"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
                    'S2
                    If fncVarKMinStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(13).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(15).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(16).Trim) = False Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0200"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
                Case "B", "BH", "BT2", "BO", "BG", "BG1", "BG2", "BG3", "BG4", "BT2G1"
                    'S1
                    If fncStdMinStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(10).Trim) = False Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0200"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
                    'S2
                    If fncStdMinStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(13).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(15).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(16).Trim) = False Then
                        intKtbnStrcSeqNo = 13
                        strMessageCd = "W0200"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
                Case "W", "WH", "WT2", "WG", "WG1", "WG2", "WG3", "WG4", "WT2G1"
                    'S1
                    If fncStdMinStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                            objKtbnStrc.strcSelection.strOpSymbol(10).Trim) = False Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0200"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
            End Select

            '*-----<< Ⅱ．最大ストロークチェック >>-----*
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "C"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "B", "W"
                            If fncVarBWMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                                     objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                     objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                                     objKtbnStrc.strcSelection.strOpSymbol(13).Trim, _
                                                     intKtbnStrcSeqNo, _
                                                     strMessageCd) = False Then
                                fncDoubleRodBaseCheck = False
                                Exit Try
                            End If
                    End Select
                Case Else
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "B", "BH", "BT", "BT1", "BT2", "BO", "BG", "BG1", "BG2", "BG3", "BG4", "BTG1", "BT1G1", "BT2G1", _
                             "BK", "BKH", "BKT", "BKT1", "BKT2", "BKO", "BKG", "BKG1", "BKG4", "BKTG1", "BKT1G1", "BKT2G1"
                            If fncVarBMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                                     objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                     objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                                     objKtbnStrc.strcSelection.strOpSymbol(13).Trim, _
                                                     intKtbnStrcSeqNo, _
                                                     strMessageCd) = False Then
                                fncDoubleRodBaseCheck = False
                                Exit Try
                            End If
                        Case "W", "WH", "WT", "WT1", "WT2", "WG", "WG1", "WG2", "WG3", "WG4", "WTG1", "WT1G1", "WT2G1", _
                             "WK", "WKH", "WKT", "WKT1", "WKT2", "WKG", "WKG1", "WKG4", "WKTG1", "WKT1G1"
                            If fncVarWMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                                     objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                     objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                                     objKtbnStrc.strcSelection.strOpSymbol(13).Trim, _
                                                     intKtbnStrcSeqNo, _
                                                     strMessageCd) = False Then
                                fncDoubleRodBaseCheck = False
                                Exit Try
                            End If
                    End Select
            End Select

            '2012/07/27 オプション外チェック
            If objKtbnStrc.strcSelection.strOtherOption.Trim <> "" Then
                '支持金具90°回転(K1)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K1") >= 0 Then
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "00"
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0430"
                            fncDoubleRodBaseCheck = False
                            Exit Try
                    End Select
                End If

                '支持金具180°回転(K2)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K2") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "LB" Then
                        intKtbnStrcSeqNo = 3
                        strMessageCd = "W0440"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
                End If

                '支持金具270°回転(K3)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K3") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "LB" Then
                        intKtbnStrcSeqNo = 3
                        strMessageCd = "W0450"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
                End If

                'P5
                Select Case objKtbnStrc.strcSelection.strKeyKataban
                    Case "C"
                        If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P5") >= 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "CB" And _
                               objKtbnStrc.strcSelection.strOpSymbol(19).IndexOf("Y") < 0 Then
                                intKtbnStrcSeqNo = 19
                                strMessageCd = "W0470"
                                fncDoubleRodBaseCheck = False
                                Exit Try
                            End If
                        End If
                    Case Else
                        If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P5") >= 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "CB" And _
                               objKtbnStrc.strcSelection.strOpSymbol(18).IndexOf("Y") < 0 Then
                                intKtbnStrcSeqNo = 18
                                strMessageCd = "W0470"
                                fncDoubleRodBaseCheck = False
                                Exit Try
                            End If
                        End If
                End Select

                'M1
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("M1") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("K") >= 0 Then
                        intKtbnStrcSeqNo = 1
                        strMessageCd = "W0760"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
                End If

                'J9
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("J9") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(17).IndexOf("J") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(17).IndexOf("K") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("O") >= 0 Then
                        intKtbnStrcSeqNo = 13
                        strMessageCd = "W0770"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
                End If

                'T9
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("T9") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("O") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("H") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("T") >= 0 Then
                        intKtbnStrcSeqNo = 1
                        strMessageCd = "W0770"
                        fncDoubleRodBaseCheck = False
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
    '*  fncHighLoadBaseCheck
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
    '*【更新履歴】
    '*                                          更新日：2008/01/10      更新者：NII A.Takahashi
    '*  ・最小ストロークの変更に伴い、バリエーション毎に最小ストロークチェックをするように修正
    '********************************************************************************************
    Private Function fncHighLoadBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Try

            fncHighLoadBaseCheck = True

            'バリエーション「Q2」＋ジャバラ「J」「L」は原価積算対応
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q2") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("J") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("L") >= 0 Then
                    intKtbnStrcSeqNo = 12
                    'RM1210067 2013/02/01 Y.Tachi ローカル版との差異修正(W0710→W0720)
                    strMessageCd = "W0720"
                    fncHighLoadBaseCheck = False
                    Exit Try
                End If
            End If

            'ジャバラ「Ｊ」「Ｌ」の最大ストロークは2000
            If objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("J") >= 0 Or _
               objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("L") >= 0 Then
                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    Case "63", "80", "100"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) > 2000 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncHighLoadBaseCheck = False
                            Exit Try
                        End If
                End Select
            End If

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "DK", "DQ2K", "DKH", "DKG", "DKG1", "DKG4"
                    If fncVarKMinStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(10).Trim, _
                                             objKtbnStrc.strcSelection.strOpSymbol(11).Trim) = False Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0200"
                        fncHighLoadBaseCheck = False
                        Exit Try
                    End If
                Case Else
                    If fncStdMinStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                        objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                                        objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                                        objKtbnStrc.strcSelection.strOpSymbol(10).Trim, _
                                                        objKtbnStrc.strcSelection.strOpSymbol(11).Trim) = False Then
                        intKtbnStrcSeqNo = 7
                        strMessageCd = "W0200"
                        fncHighLoadBaseCheck = False
                        Exit Try
                    End If
            End Select

            '*-----<< Ⅱ．最大ストロークチェック >>-----*
            '基本STチェック
            If fncStdMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                    objKtbnStrc.strcSelection.strOpSymbol(7).Trim) = False Then
                intKtbnStrcSeqNo = 7
                strMessageCd = "W0200"
                fncHighLoadBaseCheck = False
                Exit Try
            End If

            'バリエーション毎のチェック(D/P)
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("D") >= 0 Or _
               objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") >= 0 Then
                If fncVarDPMaxStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                          objKtbnStrc.strcSelection.strOpSymbol(7).Trim) = False Then
                    intKtbnStrcSeqNo = 7
                    strMessageCd = "W0200"
                    fncHighLoadBaseCheck = False
                    Exit Try
                End If
            End If

            '2012/07/27 オプション外チェック
            If objKtbnStrc.strcSelection.strOtherOption.Trim <> "" Then
                '支持金具90°回転(K1)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K1") >= 0 Then
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "00"
                            intKtbnStrcSeqNo = 3
                            strMessageCd = "W0430"
                            fncHighLoadBaseCheck = False
                            Exit Try
                    End Select
                End If

                '支持金具180°回転(K2)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K2") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "LB" Then
                        intKtbnStrcSeqNo = 3
                        strMessageCd = "W0440"
                        fncHighLoadBaseCheck = False
                        Exit Try
                    End If
                End If

                '支持金具270°回転(K3)
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("K3") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "LB" Then
                        intKtbnStrcSeqNo = 3
                        strMessageCd = "W0450"
                        fncHighLoadBaseCheck = False
                        Exit Try
                    End If
                End If

                'P5
                Select Case objKtbnStrc.strcSelection.strKeyKataban
                    Case "E"
                        If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P5") >= 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "CB" And _
                               objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("Y") < 0 Then
                                intKtbnStrcSeqNo = 14
                                strMessageCd = "W0470"
                                fncHighLoadBaseCheck = False
                                Exit Try
                            End If
                        End If
                    Case Else
                        If objKtbnStrc.strcSelection.strOtherOption.IndexOf("P5") >= 0 Then
                            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "CB" And _
                               objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("Y") < 0 Then
                                intKtbnStrcSeqNo = 13
                                strMessageCd = "W0470"
                                fncHighLoadBaseCheck = False
                                Exit Try
                            End If
                        End If
                End Select

                'M1
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("M1") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("K") >= 0 Then
                        intKtbnStrcSeqNo = 1
                        strMessageCd = "W0760"
                        fncHighLoadBaseCheck = False
                        Exit Try
                    End If
                End If

                'J9
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("J9") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("J") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("K") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G") >= 0 Then
                        intKtbnStrcSeqNo = 12
                        strMessageCd = "W0770"
                        fncHighLoadBaseCheck = False
                        Exit Try
                    End If
                End If

                'T9
                If objKtbnStrc.strcSelection.strOtherOption.IndexOf("T9") >= 0 Then
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("G") >= 0 Or _
                       objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("H") >= 0 Then
                        intKtbnStrcSeqNo = 1
                        strMessageCd = "W0770"
                        fncHighLoadBaseCheck = False
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
    '*  fncStdMinStrokeCheck
    '*【処理】
    '*  最小ストロークチェック
    '*【概要】
    '*  最小ストロークをチェックする
    '*【引数】
    '*  <String>        strMountingStyle    支持形式
    '*  <String>        strBoreSize         口径
    '*  <String>        strStroke           ストローク
    '*  <String>        strSwitchKataban    スイッチ形番
    '*  <String>        strLeadWire         リード線長さ
    '*  <String>        strSwitchQty        スイッチ数
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2007/05/16      更新者：NII A.Takahashi
    '*  ・T2W/T3Wスイッチ追加に伴い、ストロークチェックロジックを修正
    '*                                          更新日：2008/01/10      更新者：NII A.Takahashi
    '*  ・最小ストロークの変更に伴い修正
    '********************************************************************************************
    Private Function fncStdMinStrokeCheck(ByVal strMountingStyle As String, _
                                          ByVal strBoreSize As String, _
                                          ByVal strStroke As String, _
                                          ByVal strSwitchKataban As String, _
                                          ByVal strLeadWire As String, _
                                          ByVal strSwitchQty As String)

        Try

            fncStdMinStrokeCheck = False

            If Len(Trim(strSwitchKataban)) = 0 Then
                If CInt(strStroke) < 1 Then
                    Exit Try
                End If
            Else
                Select Case Trim(strSwitchKataban)
                    Case "E0"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 150 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80"
                                                If CInt(strStroke) < 145 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 140 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        If CInt(strStroke) < 335 Then
                                            Exit Try
                                        End If
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        If CInt(strStroke) < 335 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        If CInt(strStroke) < 390 Then
                                            Exit Try
                                        End If
                                    Case "4"
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 150 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80"
                                                If CInt(strStroke) < 145 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 140 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "R1", "R2", "R2Y", "R3", "R3Y", _
                         "R0", "R4", "R5", "R6", "H0", "H0Y"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        If CInt(strStroke) < 10 Then
                                            Exit Try
                                        End If
                                    Case "D"
                                        If CInt(strStroke) < 20 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80", "100"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80", "100"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                If Trim(strLeadWire) = "B" Then
                                    Select Case Trim(strSwitchQty)
                                        Case "H", "R", "D"
                                            Select Case Trim(strBoreSize)
                                                Case "40", "50"
                                                    If CInt(strStroke) < 66 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 71 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 76 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 86 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case "T", "4"
                                            Select Case Trim(strBoreSize)
                                                Case "40", "50"
                                                    If CInt(strStroke) < 92 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 97 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 102 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 112 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                    End Select
                                Else
                                    Select Case Trim(strSwitchQty)
                                        Case "H", "R", "D"
                                            Select Case Trim(strBoreSize)
                                                Case "40", "50"
                                                    If CInt(strStroke) < 86 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 91 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 96 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 106 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case "T", "4"
                                            Select Case Trim(strBoreSize)
                                                Case "40", "50"
                                                    If CInt(strStroke) < 92 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 97 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 102 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 112 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                    End Select
                                End If
                            Case "TA", "TD", "TB", "TE"
                                If Trim(strLeadWire) = "B" Then
                                    Select Case Trim(strSwitchQty)
                                        Case "H", "R"
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 28 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 26 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 31 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 34 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 40 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                    End Select
                                Else
                                    Select Case Trim(strSwitchQty)
                                        Case "H", "R"
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 38 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 36 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 41 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 44 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 50 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                    End Select
                                End If
                        End Select
                    Case "T0H", "T5H"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 65 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 70 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 175 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T0V", "T5V"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 10 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 65 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 70 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 95 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 95 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 145 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T2H", "T3H"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 10 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80", "100"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 165 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T2V", "T3V"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        If CInt(strStroke) < 10 Then
                                            Exit Try
                                        End If
                                    Case "D"
                                        If CInt(strStroke) < 15 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 80 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 95 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 90 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T2YH", "T3YH", "T2JH", "T2YD", "T2YDT", "T2YDU", _
                         "T2YLH", "T3YLH", "T1H", "T2WH", "T3WH"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 10 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80", "100"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 120 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 165 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 120 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T2YV", "T3YV", "T2JV", "T2YLV", "T3YLV", "T1V", "T2WV", "T3WV"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        If CInt(strStroke) < 10 Then
                                            Exit Try
                                        End If
                                    Case "D"
                                        If CInt(strStroke) < 15 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 70 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 80 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 90 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 90 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T8H"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63"
                                                If CInt(strStroke) < 10 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 65 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 95 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 95 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 155 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T8V"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 10 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 65 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 70 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 80 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                End Select
            End If

            fncStdMinStrokeCheck = True

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncVarPRMinStrokeCheck
    '*【処理】
    '*  最小ストロークチェック
    '*【概要】
    '*  バリエーションP・Rの最小ストロークをチェックする
    '*【引数】
    '*  <String>        strMountingStyle    支持形式
    '*  <String>        strBoreSize         口径
    '*  <String>        strStroke           ストローク
    '*  <String>        strSwitchKataban    スイッチ形番
    '*  <String>        strLeadWire         リード線長さ
    '*  <String>        strSwitchQty        スイッチ数
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2007/05/16      更新者：NII A.Takahashi
    '*  ・T2W/T3Wスイッチ追加に伴い、ストロークチェックロジックを修正
    '*                                          更新日：2008/01/10      更新者：NII A.Takahashi
    '*  ・最小ストロークの変更に伴い修正
    '********************************************************************************************
    Private Function fncVarPRMinStrokeCheck(ByVal strMountingStyle As String, _
                                            ByVal strBoreSize As String, _
                                            ByVal strStroke As String, _
                                            ByVal strSwitchKataban As String, _
                                            ByVal strLeadWire As String, _
                                            ByVal strSwitchQty As String)

        Try

            fncVarPRMinStrokeCheck = False

            If Len(Trim(strSwitchKataban)) = 0 Then
                If CInt(strStroke) < 25 Then
                    Exit Try
                End If
            Else
                Select Case Trim(strSwitchKataban)
                    Case "R1", "R2", "R2Y", "R3", "R3Y", _
                         "R0", "R4", "R5", "R6", "H0", "H0Y"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        If CInt(strStroke) < 25 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80", "100"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80", "100"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                If Trim(strLeadWire) = "B" Then
                                    Select Case Trim(strSwitchQty)
                                        Case "H", "R", "D"
                                            Select Case Trim(strBoreSize)
                                                Case "40", "50"
                                                    If CInt(strStroke) < 66 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 71 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 76 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 86 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case "T", "4"
                                            Select Case Trim(strBoreSize)
                                                Case "40", "50"
                                                    If CInt(strStroke) < 92 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 97 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 102 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 112 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                    End Select
                                Else
                                    Select Case Trim(strSwitchQty)
                                        Case "H", "R", "D"
                                            Select Case Trim(strBoreSize)
                                                Case "40", "50"
                                                    If CInt(strStroke) < 86 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 91 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 96 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 106 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case "T", "4"
                                            Select Case Trim(strBoreSize)
                                                Case "40", "50"
                                                    If CInt(strStroke) < 92 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 97 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 102 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 112 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                    End Select
                                End If
                            Case "TA", "TD"
                                Select Case Trim(strSwitchQty)
                                    Case "H"
                                        If Trim(strLeadWire) = "B" Then
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 28 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 26 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 31 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 34 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 40 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        Else
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 38 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 36 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 41 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 44 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 50 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        End If
                                End Select
                            Case "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "R"
                                        If Trim(strLeadWire) = "B" Then
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 28 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 26 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 31 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 34 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 40 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        Else
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 38 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 36 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 41 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 44 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 50 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        End If
                                End Select
                        End Select
                    Case "T0H", "T5H"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        If CInt(strStroke) < 25 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 65 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 70 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 175 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T0V", "T5V"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        If CInt(strStroke) < 25 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 65 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 70 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 95 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 95 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 145 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T2H", "T3H"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        If CInt(strStroke) < 25 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 165 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T2V", "T3V"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        If CInt(strStroke) < 25 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 80 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 95 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 90 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T2YH", "T3YH", "T2JH", "T2YD", "T2YDT", "T2YDU", _
                         "T2YLH", "T3YLH", "T1H", "T2WH", "T3WH"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        If CInt(strStroke) < 25 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 165 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T2YV", "T3YV", "T2JV", "T2YLV", "T3YLV", "T1V", "T2WV", "T3WV"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        If CInt(strStroke) < 25 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 80 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 95 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 90 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T8H"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        If CInt(strStroke) < 25 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 65 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 95 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 95 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 155 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T8V"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        If CInt(strStroke) < 25 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 65 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 70 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 80 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                End Select
            End If

            fncVarPRMinStrokeCheck = True

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncVarVMinStrokeCheck
    '*【処理】
    '*  最小ストロークチェック
    '*【概要】
    '*  バリエーションV・V1・V2の最小ストロークをチェックする
    '*【引数】
    '*  <String>        strMountingStyle    支持形式
    '*  <String>        strBoreSize         口径
    '*  <String>        strStroke           ストローク
    '*  <String>        strSwitchKataban    スイッチ形番
    '*  <String>        strLeadWire         リード線長さ
    '*  <String>        strSwitchQty        スイッチ数
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2007/05/16      更新者：NII A.Takahashi
    '*  ・T2W/T3Wスイッチ追加に伴い、ストロークチェックロジックを修正
    '*                                          更新日：2008/01/10      更新者：NII A.Takahashi
    '*  ・最小ストロークの変更に伴い修正
    '********************************************************************************************
    Private Function fncVarVMinStrokeCheck(ByVal strMountingStyle As String, _
                                           ByVal strBoreSize As String, _
                                           ByVal strStroke As String, _
                                           ByVal strSwitchKataban As String, _
                                           ByVal strLeadWire As String, _
                                           ByVal strSwitchQty As String)

        Try

            fncVarVMinStrokeCheck = False

            If Len(Trim(strSwitchKataban)) = 0 Then
                If CInt(strStroke) < 50 Then
                    Exit Try
                End If
            Else
                Select Case Trim(strSwitchKataban)
                    Case "R1", "R2", "R2Y", "R3", "R3Y", _
                         "R0", "R4", "R5", "R6", "H0", "H0Y"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D", "T"
                                        If CInt(strStroke) < 50 Then
                                            Exit Try
                                        End If
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80", "100"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                If Trim(strLeadWire) = "B" Then
                                    Select Case Trim(strSwitchQty)
                                        Case "H", "R", "D"
                                            Select Case Trim(strBoreSize)
                                                Case "40", "50"
                                                    If CInt(strStroke) < 66 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 71 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 76 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 86 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case "T", "4"
                                            Select Case Trim(strBoreSize)
                                                Case "40", "50"
                                                    If CInt(strStroke) < 92 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 97 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 102 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 112 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                    End Select
                                Else
                                    Select Case Trim(strSwitchQty)
                                        Case "H", "R", "D"
                                            Select Case Trim(strBoreSize)
                                                Case "40", "50"
                                                    If CInt(strStroke) < 86 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 91 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 96 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 106 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case "T", "4"
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 92 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 92 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 97 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 102 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 112 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                    End Select
                                End If
                            Case "TA", "TD"
                                If CInt(strStroke) < 50 Then
                                    Exit Try
                                End If
                        End Select
                    Case "T0H", "T5H"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D", "T"
                                        If CInt(strStroke) < 50 Then
                                            Exit Try
                                        End If
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 65 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 70 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 175 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T0V", "T5V"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D", "T"
                                        If CInt(strStroke) < 50 Then
                                            Exit Try
                                        End If
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 65 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 70 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 95 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 95 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 145 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "63", "80", "100"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T2H", "T3H"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                If CInt(strStroke) < 50 Then
                                    Exit Try
                                End If
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 165 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 115 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T2V", "T3V"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                If CInt(strStroke) < 50 Then
                                    Exit Try
                                End If
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 80 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 95 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 90 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        If CInt(strStroke) < 50 Then
                                            Exit Try
                                        End If
                                End Select
                        End Select
                    Case "T2YH", "T3YH", "T2JH", "T2YD", "T2YDT", "T2YDU", _
                         "T2YLH", "T3YLH", "T1H", "T2WH", "T3WH"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                If CInt(strStroke) < 50 Then
                                    Exit Try
                                End If
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 120 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 165 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 120 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T2YV", "T3YV", "T2JV", "T2YLV", "T3YLV", "T1V", "T2WV", "T3WV"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                If CInt(strStroke) < 50 Then
                                    Exit Try
                                End If
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 70 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 80 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 90 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 90 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        If CInt(strStroke) < 50 Then
                                            Exit Try
                                        End If
                                End Select
                        End Select
                End Select
            End If

            fncVarVMinStrokeCheck = True

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncVarQMinStrokeCheck
    '*【処理】
    '*  最小ストロークチェック
    '*【概要】
    '*  バリエーションQの最小ストロークをチェックする
    '*【引数】
    '*  <String>        strMountingStyle    支持形式
    '*  <String>        strBoreSize         口径
    '*  <String>        strStroke           ストローク
    '*  <String>        strSwitchKataban    スイッチ形番
    '*  <String>        strLeadWire         リード線長さ
    '*  <String>        strSwitchQty        スイッチ数
    '*【戻り値】
    '*  <Boolean>
    '*【更新履歴】
    '*                                          更新日：2007/05/16      更新者：NII A.Takahashi
    '*  ・T2W/T3Wスイッチ追加に伴い、ストロークチェックロジックを修正
    '********************************************************************************************
    Private Function fncVarQMinStrokeCheck(ByVal strVariation As String, _
                                           ByVal strMountingStyle As String, _
                                           ByVal strBoreSize As String, _
                                           ByVal strStroke As String, _
                                           ByVal strSwitchKataban As String, _
                                           ByVal strLeadWire As String, _
                                           ByVal strSwitchQty As String)

        Try

            fncVarQMinStrokeCheck = False

            If Len(Trim(strSwitchKataban)) = 0 Then
                '2011/10/31 ADD RM1110033(11月VerUP:SFRTシリーズ) START--->
                'If CInt(strStroke) < 25 Then
                '    Exit Try
                'End If
                '2011/10/31 ADD RM1110033(11月VerUP:SFRTシリーズ) <---END
                '2012/10/31 ADD RM1210086(11月VerUP:SFRTシリーズ) START--->
                'SCA2-Q2 スイッチなし最小ストロークチェック
                If Trim(strVariation) = "Q2" Then
                    If CInt(strStroke) < 5 Then
                        Exit Try
                    End If
                End If
                '2012/10/31 ADD RM1210086(11月VerUP:SFRTシリーズ) <---END
            Else
                Select Case Trim(strSwitchKataban)
                    Case "R1", "R2", "R2Y", "R3", "R3Y", _
                         "R0", "R4", "R5", "R6", "H0", "H0Y"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        If CInt(strStroke) < 10 Then
                                            Exit Try
                                        End If
                                    Case "D"
                                        If CInt(strStroke) < 20 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80", "100"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80", "100"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                If Trim(strLeadWire) = "B" Then
                                    Select Case Trim(strSwitchQty)
                                        Case "H", "R", "D"
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 95 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 93 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 97 Then
                                                        Exit Try
                                                    End If
                                                Case "80", "100"
                                                    If CInt(strStroke) < 105 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case "T"
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 130 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 132 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 136 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 142 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 152 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case "4"
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 145 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 150 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 155 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 160 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 170 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                    End Select
                                Else
                                    Select Case Trim(strSwitchQty)
                                        Case "H", "R", "D"
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 114 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 112 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 116 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 124 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 134 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case "T"
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 130 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 132 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 136 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 142 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 152 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        Case "4"
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 145 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 150 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 155 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 160 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 170 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                    End Select
                                End If
                            Case "TA", "TD"
                                Select Case Trim(strSwitchQty)
                                    Case "H"
                                        If Trim(strLeadWire) = "B" Then
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 28 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 26 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 31 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 34 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 40 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        Else
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 38 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 36 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 41 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 44 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 50 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        End If
                                End Select
                            Case "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "R"
                                        If Trim(strLeadWire) = "B" Then
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 28 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 26 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 31 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 34 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 40 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        Else
                                            Select Case Trim(strBoreSize)
                                                Case "40"
                                                    If CInt(strStroke) < 38 Then
                                                        Exit Try
                                                    End If
                                                Case "50"
                                                    If CInt(strStroke) < 36 Then
                                                        Exit Try
                                                    End If
                                                Case "63"
                                                    If CInt(strStroke) < 41 Then
                                                        Exit Try
                                                    End If
                                                Case "80"
                                                    If CInt(strStroke) < 44 Then
                                                        Exit Try
                                                    End If
                                                Case "100"
                                                    If CInt(strStroke) < 50 Then
                                                        Exit Try
                                                    End If
                                            End Select
                                        End If
                                End Select
                        End Select
                    Case "T0H", "T5H"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80", "100"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 65 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 70 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 160 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 140 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 150 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 200 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 160 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 140 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 150 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                    Case "T0V", "T5V"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 10 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 65 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 70 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 160 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 120 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 120 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 170 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 160 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 130 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 145 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                        '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                    Case "T2H", "T3H", "T2JH"
                        'Case "T2H", "T3H"
                        '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 10 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80", "100"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50"
                                                '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                                                If CInt(strStroke) < 130 Then
                                                    'If CInt(strStroke) < 135 Then
                                                    '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                                                    Exit Try
                                                End If
                                            Case "63"
                                                '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                                                If CInt(strStroke) < 135 Then
                                                    'If CInt(strStroke) < 140 Then
                                                    '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                                                    Exit Try
                                                End If
                                            Case "80"
                                                '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                                                If CInt(strStroke) < 140 Then
                                                    'If CInt(strStroke) < 145 Then
                                                    '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                                                    Exit Try
                                                End If
                                            Case "100"
                                                '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                                                If CInt(strStroke) < 150 Then
                                                    'If CInt(strStroke) < 155 Then
                                                    '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                                                If CInt(strStroke) < 190 Then
                                                    'If CInt(strStroke) < 195 Then
                                                    '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                                                    Exit Try
                                                End If
                                            Case "50"
                                                '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                                                If CInt(strStroke) < 130 Then
                                                    'If CInt(strStroke) < 135 Then
                                                    '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                                                    Exit Try
                                                End If
                                            Case "63"
                                                '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                                                If CInt(strStroke) < 135 Then
                                                    'If CInt(strStroke) < 140 Then
                                                    '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                                                    Exit Try
                                                End If
                                            Case "80"
                                                '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                                                If CInt(strStroke) < 140 Then
                                                    'If CInt(strStroke) < 145 Then
                                                    '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                                                    Exit Try
                                                End If
                                            Case "100"
                                                '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                                                If CInt(strStroke) < 150 Then
                                                    'If CInt(strStroke) < 155 Then
                                                    '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                        '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                    Case "T2V", "T3V", "T2JV"
                        'Case "T2V", "T3V"
                        '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        If CInt(strStroke) < 10 Then
                                            Exit Try
                                        End If
                                    Case "D"
                                        If CInt(strStroke) < 15 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50"
                                                '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                                                If CInt(strStroke) < 100 Then
                                                    'If CInt(strStroke) < 105 Then
                                                    '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                                                    Exit Try
                                                End If
                                                '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                                            Case "63"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 110 Then
                                                    'Case "63", "80"
                                                    'If CInt(strStroke) < 115 Then
                                                    '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                                                    Exit Try
                                                End If
                                            Case "100"
                                                '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                                                If CInt(strStroke) < 120 Then
                                                    'If CInt(strStroke) < 125 Then
                                                    '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                                                If CInt(strStroke) < 160 Then
                                                    'If CInt(strStroke) < 165 Then
                                                    '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                                                If CInt(strStroke) < 110 Then
                                                    'If CInt(strStroke) < 115 Then
                                                    '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 120 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 130 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                        '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                    Case "T8H"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "80", "100"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63"
                                                If CInt(strStroke) < 10 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 65 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 150 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 120 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 130 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 140 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 190 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 175 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 140 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 145 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 155 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "63"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                            Case "50", "80"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select

                        End Select
                    Case "T8V"
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 10 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 65 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 125 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 150 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 160 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 175 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 140 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 145 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 155 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "100"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "63", "80"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select

                        End Select
                    Case "T2YH", "T3YH", "T2YD", "T2YDT", "T2YDU", _
                         "T2YLH", "T3YLH", "T1H", "T2WH", "T3WH"
                        'Case "T2YH", "T3YH", "T2JH", "T2YD", "T2YDT", "T2YDU", _
                        ' "T2YLH", "T3YLH", "T1H", "T8H", "T2WH", "T3WH"
                        '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 10 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 20 Then
                                                    Exit Try
                                                End If
                                            Case "50", "63", "80", "100"
                                                If CInt(strStroke) < 15 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 120 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 165 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 105 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 110 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 120 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 50 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 55 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 60 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                        '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) START--->
                    Case "T2YV", "T3YV", "T2YLV", "T3YLV", "T1V", "T2WV", "T3WV"
                        'Case "T2YV", "T3YV", "T2JV", "T2YLV", "T3YLV", "T1V", "T8V", "T2WV", "T3WV"
                        '2011/05/05 MOD RM1104022(5月VerUP:SCA2-Q2シリーズ 最小ストローク修正) <---END
                        Select Case Trim(strMountingStyle)
                            Case "00", "LB", "FA", "FB", "FC", _
                                 "CA", "CB"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        If CInt(strStroke) < 10 Then
                                            Exit Try
                                        End If
                                    Case "D"
                                        If CInt(strStroke) < 15 Then
                                            Exit Try
                                        End If
                                    Case "T"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 25 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40", "50", "63"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "80", "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TC", "TF"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R", "D"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 70 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 80 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 90 Then
                                                    Exit Try
                                                End If
                                        End Select
                                    Case "T", "4"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 135 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 75 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 85 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 90 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 100 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                            Case "TA", "TD", "TB", "TE"
                                Select Case Trim(strSwitchQty)
                                    Case "H", "R"
                                        Select Case Trim(strBoreSize)
                                            Case "40"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "50"
                                                If CInt(strStroke) < 30 Then
                                                    Exit Try
                                                End If
                                            Case "63"
                                                If CInt(strStroke) < 35 Then
                                                    Exit Try
                                                End If
                                            Case "80"
                                                If CInt(strStroke) < 40 Then
                                                    Exit Try
                                                End If
                                            Case "100"
                                                If CInt(strStroke) < 45 Then
                                                    Exit Try
                                                End If
                                        End Select
                                End Select
                        End Select
                End Select
            End If

            fncVarQMinStrokeCheck = True

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncVarKMinStrokeCheck
    '*【処理】
    '*  最小ストロークチェック
    '*【概要】
    '*  バリエーションKの最小ストロークをチェックする
    '*【引数】
    '*  <String>        strMountingStyle    支持形式
    '*  <String>        strBoreSize         口径
    '*  <String>        strStroke           ストローク
    '*  <String>        strSwitchKataban    スイッチ形番
    '*  <String>        strLeadWire         リード線長さ
    '*  <String>        strSwitchQty        スイッチ数
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncVarKMinStrokeCheck(ByVal strVariation As String, _
                                           ByVal strMountingStyle As String, _
                                           ByVal strBoreSize As String, _
                                           ByVal strStroke As String, _
                                           ByVal strSwitchKataban As String, _
                                           ByVal strLeadWire As String, _
                                           ByVal strSwitchQty As String)

        Try

            fncVarKMinStrokeCheck = False

            If Len(Trim(strSwitchKataban)) = 0 Then
                If CInt(strStroke) < 1 Then
                    Exit Try
                End If
            End If
            If Trim(strVariation) = "Q2K" Then
                If CInt(strStroke) < 5 Then
                    Exit Try
                End If
            End If

            fncVarKMinStrokeCheck = True

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncStdMaxStrokeCheck
    '*【処理】
    '*  最大ストロークチェック
    '*【概要】
    '*  最大ストロークをチェックする
    '*【引数】
    '*  <String>        strBoreSize         口径
    '*  <String>        strStroke           ストローク
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncStdMaxStrokeCheck(ByVal strBoreSize As String, _
                                          ByVal strStroke As String)

        Try

            fncStdMaxStrokeCheck = False

            Select Case Trim(strBoreSize)
                Case "40"
                    If CInt(strStroke) > 1600 Then
                        Exit Try
                    End If
                Case "50"
                    If CInt(strStroke) > 2000 Then
                        Exit Try
                    End If
                Case "63", "80", "100"
                    If CInt(strStroke) > 2500 Then
                        Exit Try
                    End If
            End Select

            fncStdMaxStrokeCheck = True

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncVarOUMaxStrokeCheck
    '*【処理】
    '*  最大ストロークチェック
    '*【概要】
    '*  バリエーションO・Uの最大ストロークをチェックする
    '*【引数】
    '*  <String>        strBoreSize         口径
    '*  <String>        strStroke           ストローク
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncVarOUMaxStrokeCheck(ByVal strBoreSize As String, _
                                            ByVal strStroke As String)

        Try

            fncVarOUMaxStrokeCheck = False

            Select Case Trim(strBoreSize)
                Case "40", "50", "63"
                    If CInt(strStroke) > 600 Then
                        Exit Try
                    End If
                Case "80"
                    If CInt(strStroke) > 700 Then
                        Exit Try
                    End If
                Case "100"
                    If CInt(strStroke) > 800 Then
                        Exit Try
                    End If
            End Select

            fncVarOUMaxStrokeCheck = True

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncVarDPMaxStrokeCheck
    '*【処理】
    '*  最大ストロークチェック
    '*【概要】
    '*  バリエーションD・Pの最大ストロークをチェックする
    '*【引数】
    '*  <String>        strBoreSize         口径
    '*  <String>        strStroke           ストローク
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncVarDPMaxStrokeCheck(ByVal strBoreSize As String, _
                                            ByVal strStroke As String)

        Try

            fncVarDPMaxStrokeCheck = False

            Select Case Trim(strBoreSize)
                Case "40", "50", "63", "80", "100"
                    If CInt(strStroke) > 800 Then
                        Exit Try
                    End If
            End Select

            fncVarDPMaxStrokeCheck = True

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncVarDPMaxStrokeCheck
    '*【処理】
    '*  最大ストロークチェック
    '*【概要】
    '*  バリエーションP・Rの最大ストロークをチェックする
    '*【引数】
    '*  <String>        strBoreSize         口径
    '*  <String>        strStroke           ストローク
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncVarPRMaxStrokeCheck(ByVal strBoreSize As String, _
                                            ByVal strStroke As String)

        Try

            fncVarPRMaxStrokeCheck = False

            Select Case Trim(strBoreSize)
                Case "40", "50", "63"
                    If CInt(strStroke) > 600 Then
                        Exit Try
                    End If
                Case "80"
                    If CInt(strStroke) > 700 Then
                        Exit Try
                    End If
                Case "100"
                    If CInt(strStroke) > 800 Then
                        Exit Try
                    End If
            End Select

            fncVarPRMaxStrokeCheck = True

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncVarKMaxStrokeCheck
    '*【処理】
    '*  最大ストロークチェック
    '*【概要】
    '*  バリエーションKの最大ストロークをチェックする
    '*【引数】
    '*  <String>        strBoreSize         口径
    '*  <String>        strStroke           ストローク
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncVarKMaxStrokeCheck(ByVal strBoreSize As String, _
                                           ByVal strStroke As String)

        Try

            fncVarKMaxStrokeCheck = False

            Select Case Trim(strBoreSize)
                Case "40"
                    If CInt(strStroke) > 1600 Then
                        Exit Try
                    End If
                Case "50", "63", "80", "100"
                    If CInt(strStroke) > 1900 Then
                        Exit Try
                    End If
            End Select

            fncVarKMaxStrokeCheck = True

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncVarBMaxStrokeCheck
    '*【処理】
    '*  最大ストロークチェック
    '*【概要】
    '*  バリエーションBの最大ストロークをチェックする
    '*【引数】
    '*  <String>        strVariation        バリエーション
    '*  <String>        strBoreSize         口径
    '*  <String>        strStroke1          ストローク1
    '*  <String>        strStroke2          ストローク2
    '*  <Integer>       intKtbnStrcSeqNo    形番構成順序
    '*  <String>        strMessageCd        メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncVarBMaxStrokeCheck(ByVal strVariation As String, _
                                           ByVal strBoreSize As String, _
                                           ByVal strStroke1 As String, _
                                           ByVal strStroke2 As String, _
                                           ByRef intKtbnStrcSeqNo As Integer, _
                                           ByRef strMessageCd As String)

        Try

            fncVarBMaxStrokeCheck = False

            If bolC5CheckFlg = True Then
                If InStr(1, strVariation, "K") = 0 Then
                    Select Case Trim(strBoreSize)
                        Case "40"
                            If CInt(strStroke1) + CInt(strStroke2) > 1600 Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncVarBMaxStrokeCheck = False
                                Exit Try
                            End If
                        Case "50"
                            If CInt(strStroke1) + CInt(strStroke2) > 2000 Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncVarBMaxStrokeCheck = False
                                Exit Try
                            End If
                        Case "63", "80", "100"
                            If CInt(strStroke1) + CInt(strStroke2) > 2500 Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncVarBMaxStrokeCheck = False
                                Exit Try
                            End If
                    End Select
                Else
                    Select Case Trim(strBoreSize)
                        Case "40"
                            If CInt(strStroke1) + CInt(strStroke2) > 1600 Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncVarBMaxStrokeCheck = False
                                Exit Try
                            End If
                        Case "50", "63", "80", "100"
                            If CInt(strStroke1) + CInt(strStroke2) > 1900 Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncVarBMaxStrokeCheck = False
                                Exit Try
                            End If
                    End Select
                End If
            Else
                'S1
                Select Case Trim(strBoreSize)
                    Case "40", "50", "63"
                        If CInt(strStroke1) > 600 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncVarBMaxStrokeCheck = False
                            Exit Try
                        End If
                    Case "80"
                        If CInt(strStroke1) > 700 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncVarBMaxStrokeCheck = False
                            Exit Try
                        End If
                    Case "100"
                        If CInt(strStroke1) > 800 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncVarBMaxStrokeCheck = False
                            Exit Try
                        End If
                End Select
                'S2
                Select Case Trim(strBoreSize)
                    Case "40", "50", "63"
                        If CInt(strStroke2) > 600 Then
                            intKtbnStrcSeqNo = 13
                            strMessageCd = "W0200"
                            fncVarBMaxStrokeCheck = False
                            Exit Try
                        End If
                    Case "80"
                        If CInt(strStroke2) > 700 Then
                            intKtbnStrcSeqNo = 13
                            strMessageCd = "W0200"
                            fncVarBMaxStrokeCheck = False
                            Exit Try
                        End If
                    Case "100"
                        If CInt(strStroke2) > 800 Then
                            intKtbnStrcSeqNo = 13
                            strMessageCd = "W0200"
                            fncVarBMaxStrokeCheck = False
                            Exit Try
                        End If
                End Select
            End If

            fncVarBMaxStrokeCheck = True

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncVarBMaxStrokeCheck
    '*【処理】
    '*  最大ストロークチェック
    '*【概要】
    '*  バリエーションB,Wの最大ストロークをチェックする
    '*【引数】
    '*  <String>        strVariation        バリエーション
    '*  <String>        strBoreSize         口径
    '*  <String>        strStroke1          ストローク1
    '*  <String>        strStroke2          ストローク2
    '*  <Integer>       intKtbnStrcSeqNo    形番構成順序
    '*  <String>        strMessageCd        メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncVarBWMaxStrokeCheck(ByVal strVariation As String, _
                                           ByVal strBoreSize As String, _
                                           ByVal strStroke1 As String, _
                                           ByVal strStroke2 As String, _
                                           ByRef intKtbnStrcSeqNo As Integer, _
                                           ByRef strMessageCd As String)

        Try

            fncVarBWMaxStrokeCheck = False

            If bolC5CheckFlg = True Then
                Select Case Trim(strBoreSize)
                    Case "40", "50", "63"
                        If CInt(strStroke1) + CInt(strStroke2) > 600 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncVarBWMaxStrokeCheck = False
                            Exit Try
                        End If
                    Case "80"
                        If CInt(strStroke1) + CInt(strStroke2) > 700 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncVarBWMaxStrokeCheck = False
                            Exit Try
                        End If
                    Case "100"
                        If CInt(strStroke1) + CInt(strStroke2) > 800 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncVarBWMaxStrokeCheck = False
                            Exit Try
                        End If
                End Select
            Else
                'S1
                Select Case Trim(strBoreSize)
                    Case "40", "50", "63"
                        If CInt(strStroke1) > 600 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncVarBWMaxStrokeCheck = False
                            Exit Try
                        End If
                    Case "80"
                        If CInt(strStroke1) > 700 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncVarBWMaxStrokeCheck = False
                            Exit Try
                        End If
                    Case "100"
                        If CInt(strStroke1) > 800 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncVarBWMaxStrokeCheck = False
                            Exit Try
                        End If
                End Select
                'S2
                Select Case Trim(strBoreSize)
                    Case "40", "50", "63"
                        If CInt(strStroke2) > 600 Then
                            intKtbnStrcSeqNo = 13
                            strMessageCd = "W0200"
                            fncVarBWMaxStrokeCheck = False
                            Exit Try
                        End If
                    Case "80"
                        If CInt(strStroke2) > 700 Then
                            intKtbnStrcSeqNo = 13
                            strMessageCd = "W0200"
                            fncVarBWMaxStrokeCheck = False
                            Exit Try
                        End If
                    Case "100"
                        If CInt(strStroke2) > 800 Then
                            intKtbnStrcSeqNo = 13
                            strMessageCd = "W0200"
                            fncVarBWMaxStrokeCheck = False
                            Exit Try
                        End If
                End Select
            End If

            fncVarBWMaxStrokeCheck = True

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncVarWMaxStrokeCheck
    '*【処理】
    '*  最大ストロークチェック
    '*【概要】
    '*  バリエーションWの最大ストロークをチェックする
    '*【引数】
    '*  <String>        strVariation        バリエーション
    '*  <String>        strBoreSize         口径
    '*  <String>        strStroke1          ストローク1
    '*  <String>        strStroke2          ストローク2
    '*  <Integer>       intKtbnStrcSeqNo    形番構成順序
    '*  <String>        strMessageCd        メッセージコード
    '*【戻り値】
    '*  <Boolean>
    '*                                          更新日：2008/01/10      更新者：NII A.Takahashi
    '*  ・S1>=S2　→　S1>S2に変更
    '********************************************************************************************
    Private Function fncVarWMaxStrokeCheck(ByVal strVariation As String, _
                                           ByVal strBoreSize As String, _
                                           ByVal strStroke1 As String, _
                                           ByVal strStroke2 As String, _
                                           ByRef intKtbnStrcSeqNo As Integer, _
                                           ByRef strMessageCd As String)

        Try

            fncVarWMaxStrokeCheck = False

            If bolC5CheckFlg = True Then
                If InStr(1, strVariation, "K") = 0 Then
                    Select Case Trim(strBoreSize)
                        Case "40"
                            If CInt(strStroke1) + CInt(strStroke2) > 1600 Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncVarWMaxStrokeCheck = False
                                Exit Try
                            End If
                        Case "50"
                            If CInt(strStroke1) + CInt(strStroke2) > 2000 Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncVarWMaxStrokeCheck = False
                                Exit Try
                            End If
                        Case "63", "80", "100"
                            If CInt(strStroke1) + CInt(strStroke2) > 2500 Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncVarWMaxStrokeCheck = False
                                Exit Try
                            End If
                    End Select
                Else
                    Select Case Trim(strBoreSize)
                        Case "40"
                            If CInt(strStroke1) + CInt(strStroke2) > 1600 Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncVarWMaxStrokeCheck = False
                                Exit Try
                            End If
                        Case "50", "63", "80", "100"
                            If CInt(strStroke1) + CInt(strStroke2) > 1900 Then
                                intKtbnStrcSeqNo = 7
                                strMessageCd = "W0200"
                                fncVarWMaxStrokeCheck = False
                                Exit Try
                            End If
                    End Select
                End If
            Else
                'S1
                Select Case Trim(strBoreSize)
                    Case "40", "50", "63"
                        If CInt(strStroke1) > 600 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncVarWMaxStrokeCheck = False
                            Exit Try
                        End If
                    Case "80"
                        If CInt(strStroke1) > 700 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncVarWMaxStrokeCheck = False
                            Exit Try
                        End If
                    Case "100"
                        If CInt(strStroke1) > 800 Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncVarWMaxStrokeCheck = False
                            Exit Try
                        End If
                End Select
            End If

            'S2
            If CInt(strStroke2) > 200 Then
                intKtbnStrcSeqNo = 13
                strMessageCd = "W0200"
                fncVarWMaxStrokeCheck = False
                Exit Try
            End If

            'S1>S2
            If CInt(strStroke1) > CInt(strStroke2) Then
            Else
                intKtbnStrcSeqNo = 7
                strMessageCd = "W0610"
                fncVarWMaxStrokeCheck = False
                Exit Try
            End If

            fncVarWMaxStrokeCheck = True

        Catch ex As Exception

            Throw ex

        End Try

    End Function

End Module
