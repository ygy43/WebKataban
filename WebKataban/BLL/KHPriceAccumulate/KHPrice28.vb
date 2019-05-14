'************************************************************************************
'*  ProgramID  ：KHPrice28
'*  Program名  ：単価計算サブモジュール
'*
'*                                      作成日：2007/02/28   作成者：NII K.Sudoh
'*                                      更新日：             更新者：
'*
'*  概要       ：スーパーマイクロシリンダ　ＳＣＭ
'*
'*  ・受付No：RM0907070  二次電池対応機器　SCM
'*                                      更新日：2009/08/21   更新者：Y.Miura
'*  ・受付No：RM1001043  二次電池対応機器 チェック区分変更 3→2　
'*                                      更新日：2010/02/22   更新者：Y.Miura
'************************************************************************************
Module KHPrice28

#Region " Definition "

    Private objPrice As New KHUnitPrice
    Private bolC5Flag As Boolean
    Private strSelStrokeS1() As String = Nothing
    Private strSelStrokeS2() As String = Nothing
    Dim bolOptionP4 As Boolean                'RM0907070 2009/08/21 Y.Miura　二次電池対応

#End Region

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Try

            '配列定義
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)
            ReDim strSelStrokeS1(1)
            ReDim strSelStrokeS2(1)

            bolOptionP4 = False     'RM0907070 2009/08/21 Y.Miura　二次電池対応

            'C5チェック
            bolC5Flag = KHCylinderC5Check.fncCylinderC5Check(objKtbnStrc, False)

            '基本タイプ毎に設定
            Select Case objKtbnStrc.strcSelection.strKeyKataban
                'RM0907070 2009/08/21 Y.Miura　二次電池対応
                'Case ""
                Case "", "4", "F"
                    '基本ベース
                    Call subStandardBase(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "B", "G"
                    '背合せ・二段形ベース
                    Call subDoubleRodBase(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
                Case "D", "H"
                    '両ロッドベース
                    Call subHighLoadBase(objKtbnStrc, strOpRefKataban, decOpAmount, strPriceDiv)
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*  ProgramID  ：subStandardBase
    '*  Program名  ：基本ベース
    '************************************************************************************
    Private Sub subStandardBase(ByVal objKtbnStrc As KHKtbnStrc, _
                                ByRef strOpRefKataban() As String, _
                                ByRef decOpAmount() As Decimal, _
                                Optional ByRef strPriceDiv() As String = Nothing)

        Try

            'RM0907070 2009/08/21 Y.Miura　二次電池対応
            Dim strOpArray() As String
            Dim intLoopCnt As Integer
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(13), CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "P4", "P40"
                        bolOptionP4 = True
                End Select
            Next

            'ストローク設定
            strSelStrokeS1(0) = objKtbnStrc.strcSelection.strOpSymbol(6).Trim
            strSelStrokeS1(1) = CStr(KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                       CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                       CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)))

            '基本価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("P") < 0 Then
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "D" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "D-" & _
                                                               strSelStrokeS1(1)
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W4") < 0 Then
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        decOpAmount(UBound(decOpAmount)) = 2
                    End If
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "B-" & _
                                                               strSelStrokeS1(1)
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W4") < 0 Then
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        decOpAmount(UBound(decOpAmount)) = 2
                    End If
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If
            Else
                If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "D" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-D-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "D-" & _
                                                               strSelStrokeS1(1)
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W4") < 0 Then
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        decOpAmount(UBound(decOpAmount)) = 2
                    End If
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-D-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "B-" & _
                                                               strSelStrokeS1(1)
                    If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("W4") < 0 Then
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        decOpAmount(UBound(decOpAmount)) = 2
                    End If
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If
            End If

            'バリエーション加算価格キー
            Call subSCMVariation(objKtbnStrc, _
                                 strOpRefKataban, _
                                 decOpAmount, _
                                 strPriceDiv, _
                                 objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                 objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                 strSelStrokeS1(1))

            '支持形式加算価格キー
            Call subSCMSupport(objKtbnStrc, _
                               strOpRefKataban, _
                               decOpAmount, _
                               strPriceDiv, _
                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim, _
                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim)

            'スイッチ加算価格キー
            Call subSCMSwitch(objKtbnStrc, _
                              strOpRefKataban, _
                              decOpAmount, _
                              strPriceDiv, _
                              objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                              objKtbnStrc.strcSelection.strOpSymbol(11).Trim)

            'リード線長さ加算価格キー
            Call subSCMSwitchLead(objKtbnStrc, _
                                  strOpRefKataban, _
                                  decOpAmount, _
                                  strPriceDiv, _
                                  objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                  objKtbnStrc.strcSelection.strOpSymbol(10).Trim, _
                                  objKtbnStrc.strcSelection.strOpSymbol(11).Trim)

            'スイッチ取付け方式加算価格キー
            Call subSCMSwitchJoint(objKtbnStrc, _
                                   strOpRefKataban, _
                                   decOpAmount, _
                                   strPriceDiv, _
                                   objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                   objKtbnStrc.strcSelection.strOpSymbol(12).Trim, _
                                   objKtbnStrc.strcSelection.strOpSymbol(11).Trim, _
                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim)

            'オプション加算価格キー
            Call subSCMOption(objKtbnStrc, _
                              strOpRefKataban, _
                              decOpAmount, _
                              strPriceDiv, _
                              objKtbnStrc.strcSelection.strOpSymbol(13).Trim, _
                              objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                              strSelStrokeS1, _
                              objKtbnStrc.strcSelection.strOpSymbol(9).Trim)


            Select Case objKtbnStrc.strcSelection.strKeyKataban
                Case "F"
                    'オプション加算価格キー
                    Call subSCMOption(objKtbnStrc, _
                            strOpRefKataban, _
                            decOpAmount, _
                            strPriceDiv, _
                            objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                            objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                            strSelStrokeS1, _
                            objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If

                    '付属品加算価格キー
                    Call subSCMAccesary(objKtbnStrc, _
                             strOpRefKataban, _
                             decOpAmount, _
                             strPriceDiv, _
                             objKtbnStrc.strcSelection.strOpSymbol(15).Trim, _
                             objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                             String.Empty)
                Case Else
                    '付属品加算価格キー
                    Call subSCMAccesary(objKtbnStrc, _
                             strOpRefKataban, _
                             decOpAmount, _
                             strPriceDiv, _
                             objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                             objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                             objKtbnStrc.strcSelection.strOpSymbol(15).Trim)

                    'ロッド先端オ－ダ－メイド加算価格キー
                    Call subSCMTipOfRod(objKtbnStrc, _
                                        strOpRefKataban, _
                                        decOpAmount, _
                                        strPriceDiv, _
                                        objKtbnStrc.strcSelection.strOpSymbol(15).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim)

            End Select

         

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*  ProgramID  ：subDoubleRodBase
    '*  Program名  ：背合せ・二段形ベース
    '************************************************************************************
    Private Sub subDoubleRodBase(ByVal objKtbnStrc As KHKtbnStrc, _
                                 ByRef strOpRefKataban() As String, _
                                 ByRef decOpAmount() As Decimal, _
                                 Optional ByRef strPriceDiv() As String = Nothing)

        Try

            'ストローク設定(S1)
            strSelStrokeS1(0) = objKtbnStrc.strcSelection.strOpSymbol(6).Trim
            strSelStrokeS1(1) = CStr(KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                               CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                               CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim)))
            'ストローク設定(S2)
            strSelStrokeS2(0) = objKtbnStrc.strcSelection.strOpSymbol(12).Trim
            strSelStrokeS2(1) = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                    CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                    CInt(objKtbnStrc.strcSelection.strOpSymbol(12).Trim))

            '基本価格キー
            'S1
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "D" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "D-" & _
                                                           strSelStrokeS1(1)
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "B-" & _
                                                           strSelStrokeS1(1)
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If
            'S2
            If objKtbnStrc.strcSelection.strOpSymbol(11).Trim = "D" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "D-" & _
                                                           strSelStrokeS2(1)
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "B-" & _
                                                           strSelStrokeS2(1)
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション加算価格キー
            Call subSCMVariation(objKtbnStrc, _
                                 strOpRefKataban, _
                                 decOpAmount, _
                                 strPriceDiv, _
                                 objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                 objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                 strSelStrokeS1(1), _
                                 strSelStrokeS2(1))

            '支持形式加算価格キー
            Call subSCMSupport(objKtbnStrc, _
                               strOpRefKataban, _
                               decOpAmount, _
                               strPriceDiv, _
                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim, _
                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim)

            'スイッチ加算価格キー
            'S1
            Call subSCMSwitch(objKtbnStrc, _
                              strOpRefKataban, _
                              decOpAmount, _
                              strPriceDiv, _
                              objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                              objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
            'S2
            Call subSCMSwitch(objKtbnStrc, _
                              strOpRefKataban, _
                              decOpAmount, _
                              strPriceDiv, _
                              objKtbnStrc.strcSelection.strOpSymbol(13).Trim, _
                              objKtbnStrc.strcSelection.strOpSymbol(15).Trim)

            'リード線長さ加算価格キー
            'S1
            Call subSCMSwitchLead(objKtbnStrc, _
                                  strOpRefKataban, _
                                  decOpAmount, _
                                  strPriceDiv, _
                                  objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                  objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                                  objKtbnStrc.strcSelection.strOpSymbol(9).Trim)
            'S2
            Call subSCMSwitchLead(objKtbnStrc, _
                                  strOpRefKataban, _
                                  decOpAmount, _
                                  strPriceDiv, _
                                  objKtbnStrc.strcSelection.strOpSymbol(13).Trim, _
                                  objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                  objKtbnStrc.strcSelection.strOpSymbol(15).Trim)

            'スイッチ取付け方式加算価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(17).IndexOf("Q") < 0 Then
                'S1
                Call subSCMSwitchJoint(objKtbnStrc, _
                                       strOpRefKataban, _
                                       decOpAmount, _
                                       strPriceDiv, _
                                       objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(16).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(6).Trim)
                'S2
                Call subSCMSwitchJoint(objKtbnStrc, _
                                       strOpRefKataban, _
                                       decOpAmount, _
                                       strPriceDiv, _
                                       objKtbnStrc.strcSelection.strOpSymbol(13).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(16).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(15).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                       objKtbnStrc.strcSelection.strOpSymbol(12).Trim)
            End If

            'オプション加算価格キー
            Call subSCMOption(objKtbnStrc, _
                              strOpRefKataban, _
                              decOpAmount, _
                              strPriceDiv, _
                              objKtbnStrc.strcSelection.strOpSymbol(17).Trim, _
                              objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                              strSelStrokeS1, _
                              objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                              strSelStrokeS2, _
                              objKtbnStrc.strcSelection.strOpSymbol(13).Trim)

            Select Case objKtbnStrc.strcSelection.strKeyKataban
                '食品製造工程向け商品
                Case "G"
                    'オプション加算価格キー
                    Call subSCMOption(objKtbnStrc, _
                                      strOpRefKataban, _
                                      decOpAmount, _
                                      strPriceDiv, _
                                      objKtbnStrc.strcSelection.strOpSymbol(18).Trim, _
                                      objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                      strSelStrokeS1, _
                                      objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                      strSelStrokeS2, _
                                      objKtbnStrc.strcSelection.strOpSymbol(13).Trim)
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If

                    '付属品加算価格キー
                    Call subSCMAccesary(objKtbnStrc, _
                                        strOpRefKataban, _
                                        decOpAmount, _
                                        strPriceDiv, _
                                        objKtbnStrc.strcSelection.strOpSymbol(19).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                        String.Empty)

                Case Else
                    '付属品加算価格キー
                    Call subSCMAccesary(objKtbnStrc, _
                                        strOpRefKataban, _
                                        decOpAmount, _
                                        strPriceDiv, _
                                        objKtbnStrc.strcSelection.strOpSymbol(18).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(19).Trim)

                    'ロッド先端オ－ダ－メイド加算価格キー
                    Call subSCMTipOfRod(objKtbnStrc, _
                                        strOpRefKataban, _
                                        decOpAmount, _
                                        strPriceDiv, _
                                        objKtbnStrc.strcSelection.strOpSymbol(19).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim)
            End Select


         

        Catch ex As Exception

            Throw ex

        End Try


    End Sub

    '************************************************************************************
    '*  ProgramID  ：subHighLoadBase
    '*  Program名  ：両ロッドベース
    '************************************************************************************
    Private Sub subHighLoadBase(ByVal objKtbnStrc As KHKtbnStrc, _
                                ByRef strOpRefKataban() As String, _
                                ByRef decOpAmount() As Decimal, _
                                Optional ByRef strPriceDiv() As String = Nothing)

        Try

            'ストローク設定
            strSelStrokeS1(0) = objKtbnStrc.strcSelection.strOpSymbol(6).Trim
            strSelStrokeS1(1) = KHKataban.fncGetStrokeSize(objKtbnStrc, _
                                                    CInt(objKtbnStrc.strcSelection.strOpSymbol(3).Trim), _
                                                    CInt(objKtbnStrc.strcSelection.strOpSymbol(6).Trim))

            '基本価格キー
            If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "D" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-D-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "D-" & _
                                                           strSelStrokeS1(1)
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            Else
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-BASE-D-" & _
                                                           objKtbnStrc.strcSelection.strOpSymbol(3).Trim & "B-" & _
                                                           strSelStrokeS1(1)
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション加算価格キー
            Call subSCMVariation(objKtbnStrc, _
                                strOpRefKataban, _
                                decOpAmount, _
                                strPriceDiv, _
                                objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                strSelStrokeS1(1))

            '支持形式加算価格キー
            Call subSCMSupport(objKtbnStrc, _
                               strOpRefKataban, _
                               decOpAmount, _
                               strPriceDiv, _
                               objKtbnStrc.strcSelection.strOpSymbol(2).Trim, _
                               objKtbnStrc.strcSelection.strOpSymbol(3).Trim)

            'スイッチ加算価格キー
            Call subSCMSwitch(objKtbnStrc, _
                              strOpRefKataban, _
                              decOpAmount, _
                              strPriceDiv, _
                              objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                              objKtbnStrc.strcSelection.strOpSymbol(10).Trim)
            Call subSCMSwitchLead(objKtbnStrc, _
                                  strOpRefKataban, _
                                  decOpAmount, _
                                  strPriceDiv, _
                                  objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                                  objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                  objKtbnStrc.strcSelection.strOpSymbol(10).Trim)

            'スイッチ取付け方式加算価格キー
            Call subSCMSwitchJoint(objKtbnStrc, _
                                   strOpRefKataban, _
                                   decOpAmount, _
                                   strPriceDiv, _
                                   objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                                   objKtbnStrc.strcSelection.strOpSymbol(11).Trim, _
                                   objKtbnStrc.strcSelection.strOpSymbol(10).Trim, _
                                   objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                   objKtbnStrc.strcSelection.strOpSymbol(6).Trim)

            'オプション加算価格キー
            Call subSCMOption(objKtbnStrc, _
                              strOpRefKataban, _
                              decOpAmount, _
                              strPriceDiv, _
                              objKtbnStrc.strcSelection.strOpSymbol(12).Trim, _
                              objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                              strSelStrokeS1, _
                              objKtbnStrc.strcSelection.strOpSymbol(8).Trim)


            Select Case objKtbnStrc.strcSelection.strKeyKataban
                '食品製造工程向け商品
                Case "H"
                    'オプション加算価格キー
                    Call subSCMOption(objKtbnStrc, _
                                      strOpRefKataban, _
                                      decOpAmount, _
                                      strPriceDiv, _
                                      objKtbnStrc.strcSelection.strOpSymbol(13).Trim, _
                                      objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                      strSelStrokeS1, _
                                      objKtbnStrc.strcSelection.strOpSymbol(8).Trim)
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If

                    '付属品加算価格キー
                    Call subSCMAccesary(objKtbnStrc, _
                                        strOpRefKataban, _
                                        decOpAmount, _
                                        strPriceDiv, _
                                        objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                        String.Empty)

                Case Else
                    '付属品加算価格キー
                    Call subSCMAccesary(objKtbnStrc, _
                                        strOpRefKataban, _
                                        decOpAmount, _
                                        strPriceDiv, _
                                        objKtbnStrc.strcSelection.strOpSymbol(13).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(14).Trim)

                    'ロッド先端オ－ダ－メイド加算価格キー
                    Call subSCMTipOfRod(objKtbnStrc, _
                                        strOpRefKataban, _
                                        decOpAmount, _
                                        strPriceDiv, _
                                        objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim)
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　バリエーションによる加算を算出する
    '************************************************************************************
    Private Sub subSCMVariation(ByVal objKtbnStrc As KHKtbnStrc, _
                                ByRef strOpRefKataban() As String, _
                                ByRef decOpAmount() As Decimal, _
                                ByRef strPriceDiv() As String, _
                                ByVal strVariation As String, _
                                ByVal strBoreSize As String, _
                                ByVal strStrokeS1 As String, _
                                Optional ByVal strStrokeS2 As String = "")

        Try

            'バリエーション「X」
            If strVariation.IndexOf("X") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-X-" & _
                                                           strBoreSize
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case Else
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「Y」
            If strVariation.IndexOf("Y") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-Y-" & _
                                                           strBoreSize
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case Else
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「W4」
            If strVariation.IndexOf("W4") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-W4-" & _
                                                           strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「P」
            If strVariation.IndexOf("P") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-P-" & _
                                                           strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「R」
            If strVariation.IndexOf("R") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-R-" & _
                                                           strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「Q」
            If strVariation.IndexOf("Q") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-Q-" & _
                                                           strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「M」
            If strVariation.IndexOf("M") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-M-" & _
                                                           strBoreSize & CdCst.Sign.Hypen & strStrokeS1
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If

                If strStrokeS2.Trim <> "" Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-M-" & _
                                                               strBoreSize & CdCst.Sign.Hypen & strStrokeS2
                    decOpAmount(UBound(decOpAmount)) = 1
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If
            End If

            'バリエーション「H」
            If strVariation.IndexOf("H") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-H-" & _
                                                           strBoreSize
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case Else
                        If strVariation.IndexOf("W4") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「T」
            If strVariation.IndexOf("T") >= 0 And _
               strVariation.IndexOf("T1") < 0 And _
               strVariation.IndexOf("T2") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-T-" & _
                                                           strBoreSize
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case Else
                        If strVariation.IndexOf("W4") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「T1」
            If strVariation.IndexOf("T1") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-T1-" & _
                                                           strBoreSize
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case Else
                        If strVariation.IndexOf("W4") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「T2」
            If strVariation.IndexOf("T2") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-T2-" & _
                                                           strBoreSize
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case Else
                        If strVariation.IndexOf("W4") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「O」
            If strVariation.IndexOf("O") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-O-" & _
                                                           strBoreSize
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "B"
                        decOpAmount(UBound(decOpAmount)) = 2
                    Case Else
                        decOpAmount(UBound(decOpAmount)) = 1
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「U」
            If strVariation.IndexOf("U") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-U-" & _
                                                           strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「G」
            If strVariation.IndexOf("G") >= 0 And _
               strVariation.IndexOf("G1") < 0 And _
               strVariation.IndexOf("G2") < 0 And _
               strVariation.IndexOf("G3") < 0 And _
               strVariation.IndexOf("G4") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G-" & _
                                                           strBoreSize
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "B"
                        If strVariation.IndexOf("B") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                    Case Else
                        If strVariation.IndexOf("D") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「G1」
            If strVariation.IndexOf("G1") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G1-" & _
                                                           strBoreSize
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "B"
                        If strVariation.IndexOf("B") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                    Case Else
                        If strVariation.IndexOf("D") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「G2」
            If strVariation.IndexOf("G2") >= 0 Then
                'S1
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G2-" & _
                                                           strBoreSize & CdCst.Sign.Hypen & strStrokeS1
                If strVariation.IndexOf("D") < 0 Then
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    decOpAmount(UBound(decOpAmount)) = 2
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If

                'S2
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G2-" & _
                                                           strBoreSize & CdCst.Sign.Hypen & strStrokeS2
                If strVariation.IndexOf("D") < 0 Then
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    decOpAmount(UBound(decOpAmount)) = 2
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「G3」
            If strVariation.IndexOf("G3") >= 0 Then
                'S1
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G3-" & _
                                                           strBoreSize & CdCst.Sign.Hypen & strStrokeS1
                If strVariation.IndexOf("D") < 0 Then
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    decOpAmount(UBound(decOpAmount)) = 2
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If

                'S2
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G3-" & _
                                                           strBoreSize & CdCst.Sign.Hypen & strStrokeS2
                If strVariation.IndexOf("D") < 0 Then
                    decOpAmount(UBound(decOpAmount)) = 1
                Else
                    decOpAmount(UBound(decOpAmount)) = 2
                End If
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「G4」
            If strVariation.IndexOf("G4") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-G4-" & _
                                                           strBoreSize
                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                    Case "B"
                        If strVariation.IndexOf("B") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                    Case Else
                        If strVariation.IndexOf("D") < 0 Then
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            decOpAmount(UBound(decOpAmount)) = 2
                        End If
                End Select
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「F」
            If strVariation.IndexOf("F") >= 0 Then
                'S1
                Select Case True
                    Case CInt(strStrokeS1) <= 50
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-F-" & _
                                                                   strBoreSize & "-10-50"
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case CInt(strStrokeS1) >= 51 And CInt(strStrokeS1) <= 300
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-F-" & _
                                                                   strBoreSize & "-51-300"
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                    Case CInt(strStrokeS1) >= 301
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-F-" & _
                                                                   strBoreSize & "-301-500"
                        decOpAmount(UBound(decOpAmount)) = 1
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                End Select

                'S2
                If strStrokeS2.Trim <> "" Then
                    Select Case True
                        Case CInt(strStrokeS2) <= 50
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-F-" & _
                                                                       strBoreSize & "-10-50"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        Case CInt(strStrokeS2) >= 51 And CInt(strStrokeS2) <= 300
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-F-" & _
                                                                       strBoreSize & "-51-300"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        Case CInt(strStrokeS2) >= 301
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-F-" & _
                                                                       strBoreSize & "-301-500"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                    End Select
                End If
            End If

            'バリエーション「B」
            If strVariation.IndexOf("B") >= 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-B-" & _
                                                           strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

            'バリエーション「W」
            If strVariation.IndexOf("W") >= 0 And _
               strVariation.IndexOf("W4") < 0 Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-VAR-W-" & _
                                                           strBoreSize
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　支持形式による加算を算出する
    '************************************************************************************
    Private Sub subSCMSupport(ByVal objKtbnStrc As KHKtbnStrc, _
                              ByRef strOpRefKataban() As String, _
                              ByRef decOpAmount() As Decimal, _
                              ByRef strPriceDiv() As String, _
                              ByVal strSupport As String, _
                              ByVal strBoreSize As String)

        Try

            If strSupport.Trim <> "00" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SUPPORT-" & _
                                                           strSupport.Trim & CdCst.Sign.Hypen & strBoreSize.Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If strSupport.Trim = "LD" Then
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　スイッチによる加算を算出する
    '************************************************************************************
    Private Sub subSCMSwitch(ByVal objKtbnStrc As KHKtbnStrc, _
                             ByRef strOpRefKataban() As String, _
                             ByRef decOpAmount() As Decimal, _
                             ByRef strPriceDiv() As String, _
                             ByVal strSwitch As String, _
                             ByVal strSwitchNum As String)

        Try

            If strSwitch.Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-" & _
                                                           strSwitch.Trim
                decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwitchNum)
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　リード線の長さによる加算を算出する
    '************************************************************************************
    Private Sub subSCMSwitchLead(ByVal objKtbnStrc As KHKtbnStrc, _
                                 ByRef strOpRefKataban() As String, _
                                 ByRef decOpAmount() As Decimal, _
                                 ByRef strPriceDiv() As String, _
                                 ByVal strSwitch As String, _
                                 ByVal strSwitchLead As String, _
                                 ByVal strSwitchNum As String)

        Try

            If strSwitch.Trim <> "" Then
                If strSwitchLead.Trim <> "" Then
                    Select Case strSwitch.Trim
                        Case "T2H", "T2V", "T2YH", "T2YV", "T3H", _
                             "T3V", "T3YH", "T3YV", "T0H", "T0V", _
                             "T5H", "T5V", "T1H", "T1V", "T8H", "T8V", _
                             "T2WH", "T2WV", "T3WH", "T3WV", "T3PH", "T3PV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(1)-" & _
                                                                       strSwitchLead.Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwitchNum)
                        Case "T2YFH", "T2YFV", "T3YFH", "T3YFV", "T2YMH", _
                             "T2YMV", "T3YMH", "T3YMV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(2)-" & _
                                                                       strSwitchLead.Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwitchNum)
                        Case "T2JH", "T2JV"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(3)-" & _
                                                                       strSwitchLead.Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwitchNum)
                        Case "T2YD"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(4)-" & _
                                                                       strSwitchLead.Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwitchNum)
                        Case "T2YDT"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(5)-" & _
                                                                       strSwitchLead.Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwitchNum)
                        Case Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SWLW(6)-" & _
                                                                       strSwitchLead.Trim
                            decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwitchNum)
                    End Select
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　スイッチ取付け方式による加算を算出する
    '************************************************************************************
    Private Sub subSCMSwitchJoint(ByVal objKtbnStrc As KHKtbnStrc, _
                                  ByRef strOpRefKataban() As String, _
                                  ByRef decOpAmount() As Decimal, _
                                  ByRef strPriceDiv() As String, _
                                  ByVal strSwitch As String, _
                                  ByVal strSwitchJoint As String, _
                                  ByVal strSwitchNum As String, _
                                  ByVal strBoreSize As String, _
                                  ByVal strStroke As String)

        Try

            If strSwitch.Trim <> "" Then
                If strSwitchJoint.Trim = "" Then
                    Select Case True
                        Case CInt(strStroke) <= 300
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-JOINT-" & _
                                                                       strBoreSize.Trim & "-5-300"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        Case CInt(strStroke) >= 301 And CInt(strStroke) <= 500
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-JOINT-" & _
                                                                       strBoreSize.Trim & "-301-500"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        Case CInt(strStroke) >= 501 And CInt(strStroke) <= 1000
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-JOINT-" & _
                                                                       strBoreSize.Trim & "-501-1000"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                        Case CInt(strStroke) >= 1001
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-JOINT-" & _
                                                                       strBoreSize.Trim & "-1001-1500"
                            decOpAmount(UBound(decOpAmount)) = 1
                            If bolC5Flag = True Then
                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                            End If
                    End Select
                Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-JOINT-" & _
                                                               strSwitchJoint.Trim & CdCst.Sign.Hypen & strBoreSize.Trim
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwitchNum)
                    If bolC5Flag = True Then
                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                    End If
                End If
                'RM0907070 2009/08/21 Y.Miura　二次電池対応
                'P4加算
                If bolOptionP4 Then
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-SW-P4"
                    decOpAmount(UBound(decOpAmount)) = KHKataban.fncSwitchQtyGet(strSwitchNum)
                End If

            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　オプションによる加算を算出する
    '************************************************************************************
    Private Sub subSCMOption(ByVal objKtbnStrc As KHKtbnStrc, _
                             ByRef strOpRefKataban() As String, _
                             ByRef decOpAmount() As Decimal, _
                             ByRef strPriceDiv() As String, _
                             ByVal strOptionVar As String, _
                             ByVal strBoreSize As String, _
                             ByVal strStrokeS1() As String, _
                             ByVal strSwitchS1 As String, _
                             Optional ByVal strStrokeS2() As String = Nothing, _
                             Optional ByVal strSwitchS2 As String = "")

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            'オプション分解
            strOpArray = Split(strOptionVar, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "Q"
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "B"
                                Select Case True
                                    Case strSwitchS1 <> "" And strSwitchS2 = ""
                                        'S1
                                        Select Case True
                                            Case CInt(strStrokeS1(0)) <= 300
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-10-300"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                            Case CInt(strStrokeS1(0)) >= 301 And CInt(strStrokeS2(0)) <= 500
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-301-500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                            Case CInt(strStrokeS1(0)) >= 501 And CInt(strStrokeS2(0)) <= 1000
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-501-1000"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                            Case CInt(strStrokeS1(0)) >= 1001
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-1000-1500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                        End Select
                                    Case strSwitchS1 = "" And strSwitchS2 <> ""
                                        'S2
                                        Select Case True
                                            Case CInt(strStrokeS2(0)) <= 300
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-10-300"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                            Case CInt(strStrokeS2(0)) >= 301 And CInt(strStrokeS2(0)) <= 500
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-301-500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                            Case CInt(strStrokeS2(0)) >= 501 And CInt(strStrokeS2(0)) <= 1000
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-501-1000"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                            Case CInt(strStrokeS2(0)) >= 1001
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-1000-1500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                        End Select
                                    Case Else
                                        'S1
                                        Select Case True
                                            Case CInt(strStrokeS1(0)) <= 300
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-10-300"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                            Case CInt(strStrokeS1(0)) >= 301 And CInt(strStrokeS2(0)) <= 500
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-301-500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                            Case CInt(strStrokeS1(0)) >= 501 And CInt(strStrokeS2(0)) <= 1000
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-501-1000"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                            Case CInt(strStrokeS1(0)) >= 1001
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-1000-1500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                        End Select

                                        'S2
                                        Select Case True
                                            Case CInt(strStrokeS2(0)) <= 300
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-10-300"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                            Case CInt(strStrokeS2(0)) >= 301 And CInt(strStrokeS2(0)) <= 500
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-301-500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                            Case CInt(strStrokeS2(0)) >= 501 And CInt(strStrokeS2(0)) <= 1000
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-501-1000"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                            Case CInt(strStrokeS2(0)) >= 1001
                                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                           strBoreSize.Trim & "-1001-1500"
                                                decOpAmount(UBound(decOpAmount)) = 1
                                                If bolC5Flag = True Then
                                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                                End If
                                        End Select
                                End Select
                            Case Else
                                'スイッチ選択無しの時のみ加算
                                If strSwitchS1 = "" Then
                                    Select Case True
                                        Case CInt(strStrokeS1(0)) <= 300
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                       strBoreSize.Trim & "-10-300"
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(strStrokeS1(0)) >= 301 And CInt(strStrokeS1(0)) <= 500
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                       strBoreSize.Trim & "-301-500"
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(strStrokeS1(0)) >= 501 And CInt(strStrokeS1(0)) <= 1000
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                       strBoreSize.Trim & "-501-1000"
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                        Case CInt(strStrokeS1(0)) >= 1001
                                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                            ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                            strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-Q-" & _
                                                                                       strBoreSize.Trim & "-1001-1500"
                                            If bolC5Flag = True Then
                                                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                            End If
                                    End Select

                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "D", "H"
                                            decOpAmount(UBound(decOpAmount)) = 2
                                        Case Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                End If
                        End Select
                    Case "J", "K", "L"
                        'S1
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize.Trim & CdCst.Sign.Hypen & strStrokeS1(1)
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            Case "D", "H"
                                decOpAmount(UBound(decOpAmount)) = 2
                            Case Else
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If

                        'S2
                        If strStrokeS2 IsNot Nothing Then
                            If strStrokeS2(1) <> "" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize.Trim & CdCst.Sign.Hypen & strStrokeS2(1)
                                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                    Case "D", "H"
                                        decOpAmount(UBound(decOpAmount)) = 2
                                    Case Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            End If
                        End If
                    Case "M"
                            'S1
                            If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("M") >= 0 And _
                               strBoreSize.Trim = "32" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP*-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize.Trim & CdCst.Sign.Hypen & strStrokeS1(1)
                                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                Case "D", "H"
                                    decOpAmount(UBound(decOpAmount)) = 2
                                    Case Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize.Trim & CdCst.Sign.Hypen & strStrokeS1(1)
                                Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                Case "D", "H"
                                    decOpAmount(UBound(decOpAmount)) = 2
                                    Case Else
                                        decOpAmount(UBound(decOpAmount)) = 1
                                End Select
                                If bolC5Flag = True Then
                                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                End If
                            End If

                            'S2
                        If strStrokeS2 IsNot Nothing Then
                            If strStrokeS2(1) <> "" Then
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("M") >= 0 And _
                                   strBoreSize.Trim = "32" Then
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP*-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize.Trim & CdCst.Sign.Hypen & strStrokeS2(1)
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "D", "H"
                                            decOpAmount(UBound(decOpAmount)) = 2
                                        Case Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                    End If
                                Else
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                               strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize.Trim & CdCst.Sign.Hypen & strStrokeS2(1)
                                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                                        Case "D", "H"
                                            decOpAmount(UBound(decOpAmount)) = 2
                                        Case Else
                                            decOpAmount(UBound(decOpAmount)) = 1
                                    End Select
                                    If bolC5Flag = True Then
                                        strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                                    End If
                                End If
                            End If
                        End If
                    Case "A2"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize.Trim
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            'RM0907070 2009/08/21 Y.Miura　二次電池対応
                            'Case ""
                            Case "", "4"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "B"
                                decOpAmount(UBound(decOpAmount)) = 2
                            Case Else
                                If objKtbnStrc.strcSelection.strFullKataban.IndexOf("N13-N11") < 0 Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                        'RM0907070 2009/08/21 Y.Miura　二次電池対応
                    Case "P4", "P40"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize.Trim
                        decOpAmount(UBound(decOpAmount)) = 1

                    Case "FP1"
                        '食品製造工程向け商品
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                            Case "W4", "B", "W"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize.Trim & _
                                                                            CdCst.Sign.Hypen & objKtbnStrc.strcSelection.strOpSymbol(1).Trim

                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                           strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize.Trim
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select

                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-OP-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize.Trim
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            'RM0907070 2009/08/21 Y.Miura　二次電池対応
                            'Case ""
                            Case "", "4", "F"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case Else
                                decOpAmount(UBound(decOpAmount)) = 2
                        End Select
                        If bolC5Flag = True Then
                            strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                        End If
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　付属品による加算を算出する
    '************************************************************************************
    Private Sub subSCMAccesary(ByVal objKtbnStrc As KHKtbnStrc, _
                               ByRef strOpRefKataban() As String, _
                               ByRef decOpAmount() As Decimal, _
                               ByRef strPriceDiv() As String, _
                               ByVal strAccesary As String, _
                               ByVal strBoreSize As String, _
                               ByVal strTipOfRod As String)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Try

            strOpArray = Split(strAccesary, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "IY"
                        'I加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                   Left(strOpArray(intLoopCnt).Trim, 1) & CdCst.Sign.Hypen & strBoreSize
                        decOpAmount(UBound(decOpAmount)) = 1

                        'Y加算
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                   Right(strOpArray(intLoopCnt).Trim, 1) & CdCst.Sign.Hypen & strBoreSize
                        decOpAmount(UBound(decOpAmount)) = 1
                    Case "I", "Y"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            'RM0907070 2009/08/21 Y.Miura　二次電池対応
                            'Case ""
                            Case "", "4", "F"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "B", "G"
                                If objKtbnStrc.strcSelection.strOpSymbol(1).IndexOf("B") < 0 Then
                                    decOpAmount(UBound(decOpAmount)) = 1
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 2
                                End If
                            Case "D", "H"
                                If strTipOfRod.Trim = "" Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                    Case Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-ACC-" & _
                                                                   strOpArray(intLoopCnt).Trim & CdCst.Sign.Hypen & strBoreSize
                        Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                            'RM0907070 2009/08/21 Y.Miura　二次電池対応
                            'Case ""
                            Case "", "4"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "B", "G"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "D", "H"
                                If strTipOfRod.Trim = "" Then
                                    decOpAmount(UBound(decOpAmount)) = 2
                                Else
                                    decOpAmount(UBound(decOpAmount)) = 1
                                End If
                        End Select
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

    '************************************************************************************
    '*　ロッド先端オ－ダ－メイドによる加算を算出する
    '************************************************************************************
    Private Sub subSCMTipOfRod(ByVal objKtbnStrc As KHKtbnStrc, _
                               ByRef strOpRefKataban() As String, _
                               ByRef decOpAmount() As Decimal, _
                               ByRef strPriceDiv() As String, _
                               ByVal strTipOfRod As String, _
                               ByVal strBoreSize As String)

        Try

            If strTipOfRod.Trim <> "" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = Left(objKtbnStrc.strcSelection.strSeriesKataban.Trim, 3) & "-TIP-OF-ROD-" & _
                                                           strBoreSize.Trim
                decOpAmount(UBound(decOpAmount)) = 1
                If bolC5Flag = True Then
                    strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.C5
                End If
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
