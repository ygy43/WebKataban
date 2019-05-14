Imports Microsoft.VisualBasic
Imports WebKataban.CdCst
Imports WebKataban.ClsCommon

Public Class ClsInputCheck_17

    Public Shared intPosRowCnt As Integer = 5
    Public Shared intColCnt As Integer = 5

    Public Shared Function fncInputChk(objKtbnStrc As KHKtbnStrc, ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        fncInputChk = False
        Try
            '入力チェック2
            If Not fncInpCheck1(objKtbnStrc, strMsg, strMsgCd) Then
                '画面作成
                Exit Function
            End If
            fncInputChk = True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try
    End Function

    '********************************************************************************************
    '*【関数名】

    '*   fncInpCheck
    '*【処理】

    '*   入力チェック
    '*【引数】

    '*   strKataValues  : 形番の選択値配列          strUseValues    : 使用数の入力値配列
    '*   arySelectInf   : 設置位置の選択値配列      arySelectInf    ：設置位置の選択値配列(中間行)
    '********************************************************************************************
    Public Shared Function fncInpCheck1(objKtbnStrc As KHKtbnStrc, ByRef strMsg As String, ByRef strMsgCd As String) As Boolean
        fncInpCheck1 = False
        Dim intCnt As Integer = 0
        Try
            For intCI As Integer = 0 To intColCnt - 1
                intCnt = 0
                For intRI As Integer = Siyou_17.Unit1 - 1 To Siyou_17.Unit5 - 1
                    If objKtbnStrc.strcSelection.strPositionInfo(intRI)(intCI) = "1" Then
                        intCnt = intCnt + 1
                    End If
                Next

                '設置位置重複チェック
                If intCnt > 1 Then
                    strMsgCd = "W2510"
                    Exit Try
                ElseIf intCnt = 0 Then
                    '未接続位置チェック
                    If intCI < CInt(objKtbnStrc.strcSelection.strOpSymbol(6).ToString) Then
                        strMsgCd = "W1020"
                        strMsg = CStr(intCI + 1)
                        Exit Try
                    End If
                Else
                    If intCI >= CInt(objKtbnStrc.strcSelection.strOpSymbol(6).ToString) Then
                        strMsgCd = "W2500"
                        Exit Try
                    End If
                End If
            Next

            'ベースチェック
            If objKtbnStrc.strcSelection.strOptionKataban(Siyou_17.Base - 1).ToString.Length <= 0 Or _
                objKtbnStrc.strcSelection.intQuantity(Siyou_17.Base - 1) <> "1" Then
                strMsgCd = "W8960"
                Exit Try
            End If

            fncInpCheck1 = True
        Catch ex As Exception
            strMsg = ex.Message
            strMsgCd = "E001"
            WriteErrorLog(strMsgCd, ex)
        End Try
    End Function

End Class
