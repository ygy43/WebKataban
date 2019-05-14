'************************************************************************************
'*  ProgramID  ：KHPriceQ4
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2008/12/25   作成者：T.Sato
'*                                      更新日：             更新者：
'*
'*  概要       ：インデックスマン
'*             ：ＲＧＩＳ
'*             ：ＲＧＯＳ
'*             ：ＲＧＣＳ
'*             ：ＲＧＩＬ
'*             ：ＲＧＯＬ
'*             ：ＲＧＩＴ
'*             ：ＲＧＣＴ
'*             ：ＰＣＩＳ
'*             ：ＰＣＯＳ
'************************************************************************************
Module KHPriceQ4

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Dim intLoopCnt As Integer
        Dim intDistBetShafts As Integer     '軸間距離
        Dim intInsPosHouseMat As Integer    '据付姿勢・ハウジング材質
        Dim intLeftInSpecCD As Integer      '左入力軸仕様コード
        Dim intRightInSpecCD As Integer     '右入力軸仕様コード
        Dim intOutSpecCD As Integer         '出力軸仕様コード
        Dim intReducerSize As Integer       '減速機サイズ
        Dim intReductionRatio As Integer    '減速比
        Dim intClutchBrake As Integer       'クラッチブレーキ有無
        Dim intMotorType As Integer         'モータ種類
        Dim intMotorOutput As Integer       'モータ出力
        Dim intDrivingMethod As Integer     '駆動方法
        Dim intTsfTxgSize As Integer        'TSF・TXGサイズ
        Dim intRelTripTrqRange As Integer   'リリース・トリップトルク範囲


        Dim wkInsPosHouseMat As String    '据付姿勢・ハウジング材質

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '共通オプション位置設定
            intDistBetShafts = 1           '軸間距離
            intInsPosHouseMat = 7          '据付姿勢
            intLeftInSpecCD = 8            '左入力軸仕様コード
            intRightInSpecCD = 9           '右入力軸仕様コード
            intOutSpecCD = 10              '出力軸仕様コード

            '共通オプション変換
            ''軸間距離
            Dim strDistBetShafts As String = ""
            If objKtbnStrc.strcSelection.strOpSymbol.Length > intDistBetShafts Then
                strDistBetShafts = objKtbnStrc.strcSelection.strOpSymbol(intDistBetShafts)
            End If
            ''据付姿勢・ハウジング材質
            Dim strInsPosHouseMat As String = ""
            If objKtbnStrc.strcSelection.strOpSymbol.Length > intInsPosHouseMat Then
                strInsPosHouseMat = objKtbnStrc.strcSelection.strOpSymbol(intInsPosHouseMat)
            End If
            ''左入力軸仕様コード
            Dim strLeftSpecCD As String = ""
            If objKtbnStrc.strcSelection.strOpSymbol.Length > intLeftInSpecCD Then
                strLeftSpecCD = objKtbnStrc.strcSelection.strOpSymbol(intLeftInSpecCD)
            End If
            ''右入力軸仕様コード
            Dim strRightSpecCD As String = ""
            If objKtbnStrc.strcSelection.strOpSymbol.Length > intRightInSpecCD Then
                strRightSpecCD = objKtbnStrc.strcSelection.strOpSymbol(intRightInSpecCD)
            End If
            ''出力軸仕様コード
            Dim strOutSpecCD As String = ""
            If objKtbnStrc.strcSelection.strOpSymbol.Length > intOutSpecCD Then
                strOutSpecCD = objKtbnStrc.strcSelection.strOpSymbol(intOutSpecCD)
            End If

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                'ローラーギアカムユニット　小形/スタンダードタイプ　インデックスシリーズ、オシレートシリーズ
                Case "RGIS", "RGOS"
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        '小形タイプ
                        Case "A", "B"
                            intReducerSize = 11            '減速機サイズ
                            intReductionRatio = 12         '減速比
                            intClutchBrake = 13            'クラッチブレーキ有無
                            intMotorType = 15              'モータ種類
                            intMotorOutput = 16            'モータ出力
                            intDrivingMethod = 19          '駆動方法
                            intTsfTxgSize = 20             'TSF/TXGサイズ
                            intRelTripTrqRange = 21        'リリース・トリップトルク範囲

                            'スタンダードタイプ
                        Case "X", "Y"
                            Select Case True
                                Case strLeftSpecCD.Trim = "W", strRightSpecCD.Trim = "W"
                                    intReducerSize = 11        'ＨＯ減速機サイズ
                                Case strLeftSpecCD.Trim = "E", strRightSpecCD.Trim = "E"
                                    intReducerSize = 12        'ＴＥ減速機サイズ
                            End Select
                            intReductionRatio = 13         '減速比
                            intClutchBrake = 14            'クラッチブレーキ有無
                            intMotorType = 16              'モータ種類
                            intMotorOutput = 17            'モータ出力
                            intDrivingMethod = 20          '駆動方法
                            intTsfTxgSize = 21             'TSF/TXGサイズ
                            intRelTripTrqRange = 22        'リリース・トリップトルク範囲
                    End Select

                    'ローラーギアカムユニット　広角度タイプ　インデックスシリーズ・オシレートシリーズ
                Case "RGIL", "RGOL"
                    Select Case True
                        Case strLeftSpecCD.Trim = "W", strRightSpecCD.Trim = "W"
                            intReducerSize = 11        'ＨＯ減速機サイズ
                        Case strLeftSpecCD.Trim = "E", strRightSpecCD.Trim = "E"
                            intReducerSize = 12        'ＴＥ減速機サイズ
                    End Select
                    intReductionRatio = 13         '減速比
                    intClutchBrake = 14            'クラッチブレーキ有無
                    intTsfTxgSize = 16             'TSF/TXGサイズ
                    intRelTripTrqRange = 17        'リリース・トリップトルク範囲


                    'ローラーギアカムユニット　スタンダードタイプ　レデューサーシリーズ/ローラーギアカムユニット　テーブルタイプ　インデックスシリーズ・レデューサーシリーズ
                Case "RGCS", "RGIT", "RGCT"
                    Select Case True
                        Case strLeftSpecCD.Trim = "W", strRightSpecCD.Trim = "W"
                            intReducerSize = 11        'ＨＯ減速機サイズ
                        Case strLeftSpecCD.Trim = "E", strRightSpecCD.Trim = "E"
                            intReducerSize = 12        'ＴＥ減速機サイズ
                    End Select
                    intReductionRatio = 13         '減速比
                    intClutchBrake = 14            'クラッチブレーキ有無
                    intMotorType = 16              'モータ種類
                    intMotorOutput = 17            'モータ出力
                    intDrivingMethod = 20          '駆動方法
                    intTsfTxgSize = 21             'TSF/TXGサイズ
                    intRelTripTrqRange = 22        'リリース・トリップトルク範囲

                    'パラレルカムユニット　スタンダードタイプ　インデックスシリーズ・オシレートシリーズ
                Case "PCIS", "PCOS"
                    Select Case True
                        Case strLeftSpecCD.Trim = "W", strRightSpecCD.Trim = "W" Or _
                             strLeftSpecCD.Trim = "V", strRightSpecCD.Trim = "V"
                            intReducerSize = 11        'ＨＯ減速機サイズ
                        Case strLeftSpecCD.Trim = "E", strRightSpecCD.Trim = "E" Or _
                             strLeftSpecCD.Trim = "L", strRightSpecCD.Trim = "L"
                            intReducerSize = 12        'ＴＥ減速機サイズ
                    End Select
                    intReductionRatio = 13         '減速比
                    intClutchBrake = 14            'クラッチブレーキ有無
                    intTsfTxgSize = 16             'TSF/TXGサイズ
                    intRelTripTrqRange = 17        'リリース・トリップトルク範囲
            End Select

            'オプション変換
            ''減速機サイズ
            Dim strReducerSize As String = ""
            If objKtbnStrc.strcSelection.strOpSymbol.Length > intReducerSize Then
                strReducerSize = objKtbnStrc.strcSelection.strOpSymbol(intReducerSize)
            End If
            ''減速比
            Dim strReductionRatio As String = ""
            If objKtbnStrc.strcSelection.strOpSymbol.Length > intReductionRatio Then
                strReductionRatio = objKtbnStrc.strcSelection.strOpSymbol(intReductionRatio)
            End If
            ''クラッチブレーキ有無
            Dim strClutchBrake As String = ""
            If objKtbnStrc.strcSelection.strOpSymbol.Length > intClutchBrake Then
                strClutchBrake = objKtbnStrc.strcSelection.strOpSymbol(intClutchBrake)
            End If
            ''モータ種類
            Dim strMotorType As String = ""
            If objKtbnStrc.strcSelection.strOpSymbol.Length > intMotorType Then
                strMotorType = objKtbnStrc.strcSelection.strOpSymbol(intMotorType)
            End If
            ''モータ出力
            Dim strMotorOutput As String = ""
            If objKtbnStrc.strcSelection.strOpSymbol.Length > intMotorOutput Then
                strMotorOutput = objKtbnStrc.strcSelection.strOpSymbol(intMotorOutput)
            End If
            ''駆動方法
            Dim strDrivingMethod As String = ""
            If objKtbnStrc.strcSelection.strOpSymbol.Length > intDrivingMethod Then
                strDrivingMethod = objKtbnStrc.strcSelection.strOpSymbol(intDrivingMethod)
            End If
            ''TSF・TXGサイズ
            Dim strTsfTxgSize As String = ""
            If objKtbnStrc.strcSelection.strOpSymbol.Length > intTsfTxgSize Then
                strTsfTxgSize = objKtbnStrc.strcSelection.strOpSymbol(intTsfTxgSize)
            End If

            '基本価格キー
            If Left(strInsPosHouseMat, 1) >= "0" And Left(strInsPosHouseMat, 1) <= "9" Then
                wkInsPosHouseMat = "FC"
            Else
                wkInsPosHouseMat = "AL"
            End If

            Select Case True
                Case strLeftSpecCD = "W" Or strRightSpecCD = "W" Or _
                     strLeftSpecCD = "V" Or strRightSpecCD = "V"

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                               "*" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                               "-W" & strDistBetShafts & _
                                                               "-" & wkInsPosHouseMat & "-" & _
                                                               strReducerSize & _
                                                               strReductionRatio & _
                                                               strClutchBrake
                    decOpAmount(UBound(decOpAmount)) = 1

                Case strLeftSpecCD = "E" Or strRightSpecCD = "E" Or _
                     strLeftSpecCD = "L" Or strRightSpecCD = "L"

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                               "*" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                               "-E" & strDistBetShafts & _
                                                               "-" & wkInsPosHouseMat & "-" & _
                                                               strReducerSize & _
                                                               strReductionRatio & _
                                                               strClutchBrake
                    decOpAmount(UBound(decOpAmount)) = 1

                Case strLeftSpecCD = "G" Or strRightSpecCD = "G"

                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                               "*" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                               "-G" & strDistBetShafts & _
                                                               "-" & wkInsPosHouseMat & "-" & _
                                                               strMotorType & _
                                                               strMotorOutput & _
                                                               strDrivingMethod
                    decOpAmount(UBound(decOpAmount)) = 1

                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)

                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                               "*" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                               strDistBetShafts & "-" & wkInsPosHouseMat
                    decOpAmount(UBound(decOpAmount)) = 1

            End Select

            '入出力仕様加算価格キー
            For intLoopCnt = intLeftInSpecCD To intOutSpecCD

                If intLoopCnt >= objKtbnStrc.strcSelection.strOpSymbol.Length Then
                    Exit For
                End If

                If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt) <> "" And _
                   objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt) <> "N" And _
                   (intLoopCnt <> 9 Or strLeftSpecCD <> "K" Or strRightSpecCD <> "K") Then

                    Select Case True
                        Case intLoopCnt = intOutSpecCD

                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                       "*" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                                       strDistBetShafts & "-O" & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt)
                            decOpAmount(UBound(decOpAmount)) = 1

                        Case Else
                            If objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt) = "H" Then

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                'H中空軸型ｷﾞﾔｰﾄﾞﾓｰﾀ付時、ﾓｰﾀ出力記号を付加
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                           "*" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                                           strDistBetShafts & "-I" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt) & _
                                                                           strMotorOutput
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                                           "*" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                                           strDistBetShafts & "-I" & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(intLoopCnt)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                    End Select
                End If
            Next

            'オプション加算価格キー
            If Len(Trim(strTsfTxgSize)) <> 0 Then

                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 2) & _
                                                           "*" & Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1) & _
                                                           strDistBetShafts

                Select Case strOutSpecCD
                    Case "F", "A", "S", "B"
                        Select Case Mid(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 4, 1)
                            Case "S", "L"       'RG*S/RG*L/PC*S
                                '2010/09/15 MOD RM1009006(10月VerUP:インデックスマン対応) START--->
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & _
                                                                           "-TSF" & strTsfTxgSize
                                'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & _
                                '                                           "-TSF" & strTsfTxgSize
                                'decOpAmount(UBound(decOpAmount)) = 1
                                '2010/09/15 MOD RM1009006(10月VerUP:インデックスマン対応) <---END

                            Case "T"            'RG*T

                                '2010/09/15 MOD RM1009006(10月VerUP:インデックスマン対応) START--->
                                strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & _
                                                                           "-TST" & strTsfTxgSize
                                'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                'strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & _
                                '                                           "-TST" & strTsfTxgSize
                                'decOpAmount(UBound(decOpAmount)) = 1
                                '2010/09/15 MOD RM1009006(10月VerUP:インデックスマン対応) <---END

                        End Select
                    Case "X", "C", "Y", "D" 'RG*S/RG*L/PC*S

                        '2010/09/15 MOD RM1009006(10月VerUP:インデックスマン対応) START--->
                        strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & _
                                                                   "-TGX" & strTsfTxgSize
                        'ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        'strOpRefKataban(UBound(strOpRefKataban)) = strOpRefKataban(UBound(strOpRefKataban)) & _
                        '                                           "-TGX" & strTsfTxgSize
                        'decOpAmount(UBound(decOpAmount)) = 1
                        '2010/09/15 MOD RM1009006(10月VerUP:インデックスマン対応) <---END

                End Select

                '2010/09/15 ADD RM1009006(10月VerUP:インデックスマン対応) START--->
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                decOpAmount(UBound(decOpAmount)) = 1
                '2010/09/15 ADD RM1009006(10月VerUP:インデックスマン対応) <---END

            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
