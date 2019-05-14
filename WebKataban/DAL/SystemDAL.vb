Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class SystemDAL

    ''' <summary>
    ''' 稼動状況確認
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strBaseCd">拠点コード</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncOpeStateChk(ByVal objCon As SqlConnection, ByVal strBaseCd As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL作成
            sbSql.Append(" SELECT  b.operation_status ")
            sbSql.Append(" FROM    kh_base_mst  a ")
            sbSql.Append(" INNER JOIN  kh_ope_state  b ")
            sbSql.Append(" ON      a.base_cd             = b.base_cd ")
            sbSql.Append(" WHERE   a.base_cd             = @BaseCd")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@BaseCd", SqlDbType.Char, 2).Value = strBaseCd
            End With
            '実行
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dtResult)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try

        Return dtResult
    End Function
End Class
