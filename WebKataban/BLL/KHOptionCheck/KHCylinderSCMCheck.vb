Module KHCylinderSCMCheck

    '********************************************************************************************
    '*【関数名】
    '*  fncCheckSelectOption
    '*【処理】
    '*  シリンダチェック
    '*【概要】
    '*  シリンダＳＣＭシリーズをチェックする
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
                Case "SCM"
                    '基本ベース毎にチェック
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        'RM0907070 2009/08/24 Y.Miura　二次電池対応
                        'Case ""
                        Case "", "4", "F"
                            '基本ベースチェック
                            If fncStandardBaseCheck(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case "B", "G"
                            '背合わせ＆２段形ベースチェック
                            If fncDoubleRodBaseCheck(objKtbnStrc, _
                                                     intKtbnStrcSeqNo, _
                                                     strOptionSymbol, _
                                                     strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case "D", "H"
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
    '*
    '********************************************************************************************
    Private Function fncStandardBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Try

            fncStandardBaseCheck = True

            'バリエーション「Q」＋ジャバラ「J」「K」「L」は原価積算対応  RM1701061  2017/02/01 追加 松原
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("Q") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("J") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("K") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(13).IndexOf("L") >= 0 Then
                    intKtbnStrcSeqNo = 13
                    strMessageCd = "W0580"
                    fncStandardBaseCheck = False
                    Exit Try
                End If
            End If

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            'スイッチ毎のチェック
            If fncSCMSwitchStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(11).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(12).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = False Then
                intKtbnStrcSeqNo = 6
                strMessageCd = "W0200"
                fncStandardBaseCheck = False
                Exit Try
            End If

            '*-----<< Ⅱ．最大ストロークチェック >>-----*
            'バリエーション毎のチェック
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "X", "Y", "XM", "XT2", "YM", "YT2"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 200 Then
                        intKtbnStrcSeqNo = 6
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                Case "F", "RF"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 500 Then
                        intKtbnStrcSeqNo = 6
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                Case "P", "PH", "PT2", "M", "PM", "RM"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 600 Then
                        intKtbnStrcSeqNo = 6
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                Case "W4", "W4M", "W4H", "W4T", "W4T1", _
                     "W4T2", "W4G", "W4G1", "W4G2", "W4G3", _
                     "W4G4", "W4HG", "W4TG1", "W4T1G1", "W4T2G1", _
                     "W4T2G4"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "20", "25", "32"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 500 Then
                                intKtbnStrcSeqNo = 6
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        Case Else
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 600 Then
                                intKtbnStrcSeqNo = 6
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                    End Select
                Case Else
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "20", "25", "32"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 1000 Then
                                intKtbnStrcSeqNo = 6
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                    End Select
            End Select

            '支持形式毎のチェック '「LD」選択時は300mmまで
            If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "LD" Then
                If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 300 Then
                    intKtbnStrcSeqNo = 6
                    strMessageCd = "W0200"
                    fncStandardBaseCheck = False
                    Exit Try
                End If
            End If

            '*-----<< Ⅲ．付属品のチェック >>-----*
            '「B1」を選択した時、支持口径が「CB」又は、付属品「Y」でないときはエラー
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "F"
                    If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("B1") >= 0 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "CB" And _
                           objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("Y") < 0 Then
                            intKtbnStrcSeqNo = 15
                            strMessageCd = "W0290"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If

                Case Else
                    If objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("B1") >= 0 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(2).Trim <> "CB" And _
                           objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("Y") < 0 Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0290"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If
            End Select

            '「B1」を選択した時、口径が「80」又は「100」でないときはエラー
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "F"
                    If objKtbnStrc.strcSelection.strOpSymbol(15).IndexOf("B1") >= 0 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "80" And _
                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "100" Then
                            intKtbnStrcSeqNo = 15
                            strMessageCd = "W0600"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If
                Case Else
                    If objKtbnStrc.strcSelection.strOpSymbol(14).IndexOf("B1") >= 0 Then
                        If objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "80" And _
                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim <> "100" Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0600"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If
            End Select

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
    '********************************************************************************************
    Private Function fncDoubleRodBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                           ByRef intKtbnStrcSeqNo As Integer, _
                                           ByRef strOptionSymbol As String, _
                                           ByRef strMessageCd As String) As Boolean

        Try

            fncDoubleRodBaseCheck = True

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            'S1:ストローク
            If fncSCMSwitchStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(16).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = False Then
                intKtbnStrcSeqNo = 6
                strMessageCd = "W0200"
                fncDoubleRodBaseCheck = False
                Exit Try
            End If
            'S2:ストローク
            If fncSCMSwitchStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(12).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(13).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(16).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = False Then
                intKtbnStrcSeqNo = 12
                strMessageCd = "W0200"
                fncDoubleRodBaseCheck = False
                Exit Try
            End If

            '*-----<< Ⅱ．最大ストロークチェック >>-----*
            'バリエーション毎のチェック
            'S1:ストローク
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "XB", "XBT2", "YB", "YBT2"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 200 Then
                        intKtbnStrcSeqNo = 6
                        strMessageCd = "W0200"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
                Case "BF"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 500 Then
                        intKtbnStrcSeqNo = 6
                        strMessageCd = "W0200"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
                Case "W", "WM", "WH", "WT", "WT1", _
                     "WT2", "WG", "WG1", "WG2", "WG3", _
                     "WG4", "WHG", "WTG1", "WT1G1", "WT2G1", _
                     "WT2G4"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) > 600 Then
                        intKtbnStrcSeqNo = 6
                        strMessageCd = "W0200"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
            End Select
            'S2:ストローク
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "XB", "XBT2", "YB", "YBT2"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 200 Then
                        intKtbnStrcSeqNo = 12
                        strMessageCd = "W0200"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
                Case "BF"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 500 Then
                        intKtbnStrcSeqNo = 12
                        strMessageCd = "W0200"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
                Case "W", "WM", "WH", "WT", "WT1", _
                     "WT2", "WG", "WG1", "WG2", "WG3", _
                     "WG4", "WHG", "WTG1", "WT1G1", "WT2G1", _
                     "WT2G4"
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 200 Then
                        intKtbnStrcSeqNo = 12
                        strMessageCd = "W0200"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
            End Select

            '*-----<< Ⅲ．Ｓ１：Ｓ２　ストロークチェック >>-----*
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "B", "BH", "BT", "BT1", "BT2", _
                     "BO", "BG", "BG1", "BG2", "BG3", _
                     "BG4", "BHG", "BTG1", "BT1G1", "BT2G1", _
                     "BT2G4"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "20", "25", "32"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 1000 Then
                                intKtbnStrcSeqNo = 6
                                strMessageCd = "W0200"
                                fncDoubleRodBaseCheck = False
                                Exit Try
                            End If
                        Case Else
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 1500 Then
                                intKtbnStrcSeqNo = 12
                                strMessageCd = "W0200"
                                fncDoubleRodBaseCheck = False
                                Exit Try
                            End If
                    End Select
                Case "W", "WM", "WH", "WT", "WT1", _
                     "WT2", "WG", "WG1", "WG2", "WG3", _
                     "WG4", "WHG", "WTG1", "WT1G1", "WT2G1", _
                     "WT2G4"
                    Select Case True
                        Case CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) < CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim)
                            intKtbnStrcSeqNo = 6
                            strMessageCd = "W0610"
                            fncDoubleRodBaseCheck = False
                            Exit Try
                    End Select

                    Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                        Case "20", "25", "32"
                            If CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim) + CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim) > 1000 Then
                                intKtbnStrcSeqNo = 6
                                strMessageCd = "W0200"
                                fncDoubleRodBaseCheck = False
                                Exit Try
                            End If
                    End Select
            End Select

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
    '********************************************************************************************
    Private Function fncHighLoadBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Try

            fncHighLoadBaseCheck = True

            'バリエーション「DQ」＋ジャバラ「J」「K」「L」は原価積算対応  RM1701061  2017/02/01 追加 松原
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("DQ") >= 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("J") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("K") >= 0 Or _
                   objKtbnStrc.strcSelection.strOpSymbol(12).IndexOf("L") >= 0 Then
                    intKtbnStrcSeqNo = 12
                    strMessageCd = "W0580"
                    fncHighLoadBaseCheck = False
                    Exit Try
                End If
            End If

            '*-----<< Ⅰ．最小ストロークチェック >>-----*
            'スイッチ毎のチェック
            If fncSCMSwitchStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(10).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(2).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(11).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = False Then
                intKtbnStrcSeqNo = 6
                strMessageCd = "W0200"
                fncHighLoadBaseCheck = False
                Exit Try
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*【関数名】
    '*  fncSCMSwitchStrokeCheck
    '*【処理】
    '*  スイッチ毎のチェック
    '*【概要】
    '*  スイッチ形番毎にストロークをチェックする
    '*【引数】
    '*  <String>        strStroke           ストローク
    '*  <String>        strSwitchKataban    スイッチ形番
    '*  <String>        strSwitchQty        スイッチ数
    '*  <String>        strVariation        バリエーション
    '*  <String>        strSupport          支持形式
    '*  <String>        strSwitchJoint      スイッチ取付方式
    '*  <String>        strPortSize         口径
    '*【戻り値】
    '*  <Boolean>
    '*【修正履歴】
    '*                                 更新日：2007/06/26   更新者：NII A.Takahashi
    '*  　・口径を引数に設定する
    '*  　・最小ストローク変更のため修正する
    '********************************************************************************************
    Private Function fncSCMSwitchStrokeCheck(ByVal strStroke As String, _
                                             ByVal strSwitchKataban As String, _
                                             ByVal strSwitchQty As String, _
                                             ByVal strVariation As String, _
                                             ByVal strSupport As String, _
                                             ByVal strSwitchJoint As String, _
                                             ByVal strPortSize As String)

        Dim objPrice As New KHUnitPrice

        Try

            fncSCMSwitchStrokeCheck = False

            'SW選択有無判定
            If strSwitchKataban.Trim = "" Then
                'バリエーションで判定
                If strVariation.IndexOf("X") < 0 And _
                   strVariation.IndexOf("Y") < 0 Then
                    If CInt(strStroke) < 10 Then
                        Exit Try
                    End If
                Else
                    If CInt(strStroke) < 5 Then
                        Exit Try
                    End If
                End If
            Else
                'スイッチ取付け方式判定
                'レール方式
                If strSwitchJoint.Trim = "" Then
                    'スイッチ個数で判定
                    Select Case KHKataban.fncSwitchQtyGet(strSwitchQty)
                        Case "1"
                            Select Case strSupport
                                Case "TA", "TB"
                                    If CInt(strStroke) < 25 Then
                                        Exit Try
                                    End If
                                Case Else
                                    If CInt(strStroke) < 10 Then
                                        Exit Try
                                    End If
                            End Select
                        Case "2"
                            If CInt(strStroke) < 25 Then
                                Exit Try
                            End If
                        Case "3"
                            Select Case strSwitchKataban.Trim
                                Case "T0H", "T0V", "T5H", "T5V"
                                    If CInt(strStroke) < 55 Then
                                        Exit Try
                                    End If
                                Case "T2H", "T2V", "T3H", "T3V"
                                    If CInt(strStroke) < 50 Then
                                        Exit Try
                                    End If
                                Case Else
                                    Select Case strPortSize
                                        Case "20", "25", "32", "40"
                                            If CInt(strStroke) < 70 Then
                                                Exit Try
                                            End If
                                        Case "50", "63", "80", "100"
                                            If CInt(strStroke) < 65 Then
                                                Exit Try
                                            End If
                                    End Select
                            End Select
                        Case "4"
                            Select Case strSwitchKataban.Trim
                                Case "T0H", "T0V", "T2H", "T2V", "T3H", "T3V", "T5H", "T5V"
                                    If CInt(strStroke) < 55 Then
                                        Exit Try
                                    End If
                                Case Else
                                    Select Case strPortSize
                                        Case "20", "25", "32", "40"
                                            If CInt(strStroke) < 70 Then
                                                Exit Try
                                            End If
                                        Case "50", "63", "80", "100"
                                            If CInt(strStroke) < 65 Then
                                                Exit Try
                                            End If
                                    End Select
                            End Select
                        Case "5"
                            Select Case strSwitchKataban.Trim
                                Case "T0H", "T0V", "T5H", "T5V"
                                    If CInt(strStroke) < 90 Then
                                        Exit Try
                                    End If
                                Case "T2H", "T2V", "T3H", "T3V"
                                    If CInt(strStroke) < 75 Then
                                        Exit Try
                                    End If
                                Case Else
                                    If CInt(strStroke) < 110 Then
                                        Exit Try
                                    End If
                            End Select
                    End Select
                Else
                    'バンド方式
                    'スイッチ個数で判定
                    Select Case KHKataban.fncSwitchQtyGet(strSwitchQty)
                        Case "1"
                            If CInt(strStroke) < 10 Then
                                Exit Try
                            End If
                        Case "2"
                            Select Case strSwitchKataban.Trim
                                Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                                     "T3V", "T5H", "T5V"
                                    If CInt(strStroke) < 25 Then
                                        Exit Try
                                    End If
                                Case Else
                                    If CInt(strStroke) < 35 Then
                                        Exit Try
                                    End If
                            End Select
                        Case "3"
                            Select Case strSwitchKataban.Trim
                                Case "T0H", "T0V", "T2H", "T2V", "T3H", _
                                     "T3V", "T5H", "T5V"
                                    If CInt(strStroke) < 50 Then
                                        Exit Try
                                    End If
                                Case Else
                                    If CInt(strStroke) < 55 Then
                                        Exit Try
                                    End If
                            End Select
                        Case "4"
                            Select Case strSwitchKataban.Trim
                                Case "T0H", "T0V", "T5H", "T5V"
                                    If CInt(strStroke) < 70 Then
                                        Exit Try
                                    End If
                                Case "T2H", "T2V", "T3H", "T3V"
                                    If CInt(strStroke) < 75 Then
                                        Exit Try
                                    End If
                                Case Else
                                    If CInt(strStroke) < 80 Then
                                        Exit Try
                                    End If
                            End Select
                        Case "5"
                            Select Case strSwitchKataban.Trim
                                Case "T0H", "T0V", "T2H", "T2V", "T3H", "T3V", "T5H", "T5V"
                                    If CInt(strStroke) < 95 Then
                                        Exit Try
                                    End If
                                Case Else
                                    If CInt(strStroke) < 100 Then
                                        Exit Try
                                    End If
                            End Select
                    End Select
                End If
            End If

            fncSCMSwitchStrokeCheck = True

        Catch ex As Exception

            Throw ex

        Finally

            objPrice = Nothing

        End Try

    End Function

End Module
