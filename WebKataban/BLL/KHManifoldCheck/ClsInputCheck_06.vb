Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_06

    Public Shared intPosRowCnt As Integer = 18
    Public Shared intColCnt As Integer = 10

    '********************************************************************************************
    '*【関数名】
    '*   fncInpCheck
    '*【処理】
    '*   入力チェック
    '********************************************************************************************
    Public Shared Function fncInpChk(objKtbnStrc As KHKtbnStrc, _
                                     ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        Dim intMaxNo As Integer
        Dim intCount As Integer
        Dim sbCoordinates As New StringBuilder
        Dim hshtKataban As New Hashtable
        Dim strCoordinates As String = ""
        Dim bolReturn As Boolean

        Try
            Dim strUseValues() As Double = objKtbnStrc.strcSelection.intQuantity
            Dim strKataValues() As String = objKtbnStrc.strcSelection.strOptionKataban
            Dim arySelectInf() As String = objKtbnStrc.strcSelection.strPositionInfo

            '************** Ａ・Ｂポート接続口径チェック(1.2) ************************************
            If objKtbnStrc.strcSelection.strOpSymbol(2).ToString = "XX" Then
                intCount = 0

                For intLoop As Integer = 0 To CInt(objKtbnStrc.strcSelection.strOpSymbol(1)) - 1
                    intCount = 0
                    '選択状態の形番数をカウント
                    For intRI As Integer = Siyou_06.ABCon01 - 1 To Siyou_06.ABCon1Z - 1
                        If arySelectInf(intRI)(intLoop) = "1" Then
                            intCount = intCount + 1
                        End If
                    Next

                    'カウントが０の場合、エラー
                    If intCount = 0 Then
                        For intRx As Integer = Siyou_06.ABCon01 - 1 To Siyou_06.ABCon1Z - 1
                            sbCoordinates.Append(CStr(intRx + 1) & strComma & (intLoop + 1).ToString & strPipe)
                        Next
                        strMsg = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                        strMsgCd = "W2060"
                        bolReturn = False
                        Exit Try
                    End If
                Next

            End If
            sbCoordinates = New StringBuilder

            '以下の条件の場合、入力チェックをとばす
            If objKtbnStrc.strcSelection.strOpSymbol(5).ToString = "" Then
                bolReturn = True
                Exit Try
            End If

            intMaxNo = CInt(objKtbnStrc.strcSelection.strOpSymbol(1).ToString)

            '************** 形番チェック ***********************************************
            For intRI As Integer = 0 To Siyou_06.ExpCovExh - 1


                '形番が未選択の場合
                If Len(Trim(strKataValues(intRI))) = 0 Then

                    'ABポート以外の行で設置位置が選択されていたらエラー
                    If (intRI < Siyou_06.ABCon01 - 1 Or intRI > Siyou_06.ABCon1Z - 1) And Int(strUseValues(intRI)) > 0 Then
                        strMsgCd = "W1400"
                        bolReturn = False
                        Exit Try
                    End If
                Else
                    '形番重複チェック
                    If hshtKataban.ContainsKey(strKataValues(intRI)) Then
                        strMsgCd = "W1330"
                        bolReturn = False
                        Exit Try
                    Else
                        hshtKataban.Add(strKataValues(intRI), "")
                    End If
                End If
            Next

            '************** 電磁連数弁形式チェック(1.1) ************************************

            '列ごとの選択数をカウント
            For intCI As Integer = 0 To intColCnt - 1
                intCount = 0
                For intRI As Integer = Siyou_06.Elect1 - 1 To Siyou_06.Elect6 - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                        intCount = intCount + 1
                    End If
                Next

                If intCount > 1 Then
                    '一列につき２つ以上選択されていたらエラー
                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                    strMsg = strCoordinates
                    strMsgCd = "W1790"
                    bolReturn = False
                    Exit Try
                End If
                sbCoordinates = New StringBuilder
            Next

            '電磁弁選択数が連数より少ない場合、エラー
            intCount = 0
            For intRI As Integer = Siyou_06.Elect1 - 1 To Siyou_06.Elect6 - 1
                intCount = intCount + Int(strUseValues(intRI))
            Next
            If intCount < intMaxNo Then
                strMsgCd = "W1180"
                bolReturn = False
                Exit Try
            End If

            '一列につき２つ以上選択されていたらエラー
            For intCI As Integer = 0 To intColCnt - 1
                intCount = 0
                For intRI As Integer = Siyou_06.ABCon01 - 1 To Siyou_06.ABCon1Z - 1
                    If arySelectInf(intRI)(intCI) = "1" Then
                        intCount = intCount + 1
                        sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                    End If
                Next
                If intCount > 1 Then
                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                    strMsg = strCoordinates
                    strMsgCd = "W1950"
                    bolReturn = False
                    Exit Try
                End If
                sbCoordinates = New StringBuilder
            Next

            '************** 給気スペーサチェック(1.4) ************************************
            For intCI As Integer = 0 To intColCnt - 1
                intCount = 0
                For intRI As Integer = Siyou_06.RepSpace1 - 1 To Siyou_06.RepSpace2 - 1
                    If arySelectInf(intRI)(intCI) = "1" Then

                        intCount = intCount + 1
                        sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                    End If
                Next
                If intCount > 1 Then
                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                    strMsg = strCoordinates
                    strMsgCd = "W1960"
                    bolReturn = False
                    Exit Try
                End If
                sbCoordinates = New StringBuilder
            Next

            '************** 排気スペーサチェック(1.5) ************************************
            For intCI As Integer = 0 To intColCnt - 1
                intCount = 0
                For intRI As Integer = Siyou_06.ExhSpace1 - 1 To Siyou_06.ExhSpace2 - 1
                    If arySelectInf(intRI)(intCI) = "1" Then

                        intCount = intCount + 1
                        sbCoordinates.Append(CStr(intRI + 1) & strComma & CStr(intCI + 1) & strPipe)
                    End If
                Next
                If intCount > 1 Then
                    strCoordinates = Left(sbCoordinates.ToString, Len(sbCoordinates.ToString) - 1)
                    strMsg = strCoordinates
                    strMsgCd = "W1970"
                    bolReturn = False
                    Exit Try
                End If
                sbCoordinates = New StringBuilder
            Next

            bolReturn = True
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            hshtKataban.Clear()
            hshtKataban = Nothing
        End Try
        fncInpChk = bolReturn
    End Function

End Class
