Imports System.Data.SqlClient
Imports WebKataban.ClsCommon

Public Class KHSystem

    ''' <summary>
    ''' 稼動状況確認
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strBaseCd">拠点コード</param>
    ''' <returns>稼動状況     0：停止／1：稼動／E：エラー</returns>
    ''' <remarks>引数で渡された拠点コードにて稼動状況を確認する</remarks>
    Public Function fncOpeStateChk(ByVal objCon As SqlConnection, ByVal strBaseCd As String) As String
        Dim dt As New DataTable
        Dim dalSystem As New SystemDAL

        Try
            dt = dalSystem.fncOpeStateChk(objCon, strBaseCd)

            If dt.Rows.Count > 0 Then
                '戻り値に稼動状況を設定
                fncOpeStateChk = dt.Rows(0)("operation_status")
            Else
                '稼動状況が取得出来ない場合はトラブル
                fncOpeStateChk = CdCst.OpeState.Trouble
            End If

        Catch ex As Exception
            fncOpeStateChk = CdCst.OpeState.Trouble
            WriteErrorLog("E001", ex)
        End Try

    End Function

End Class
