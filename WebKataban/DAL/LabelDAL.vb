Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class LabelDAL

    ''' <summary>
    ''' ラベルデータ取り込み
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strLangCd">プログラムＩＤ</param>
    ''' <param name="strPgmId">言語コード</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function fncSelectPageAllLabels(ByVal objCon As SqlConnection, ByVal strLangCd As String, ByVal strPgmId As String) As DataTable
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand = Nothing
        Dim objAdp As SqlDataAdapter = Nothing
        fncSelectPageAllLabels = New DataTable

        Try
            sbSql.Append(" SELECT  ISNULL(b.label_content, a.label_content) AS label_content,")
            sbSql.Append(" a.label_div AS label_div, ")
            sbSql.Append(" a.label_seq AS label_seq ")
            sbSql.Append(" FROM    sales.kh_label_mst2 a ")
            sbSql.Append(" LEFT JOIN  sales.kh_label_mst2 b ")
            sbSql.Append(" ON      a.program_id  = b.program_id ")
            sbSql.Append(" AND     a.label_div   = b.label_div ")
            sbSql.Append(" AND     a.label_seq   = b.label_seq ")
            sbSql.Append(" AND     b.language_cd = @language_cd ")
            sbSql.Append(" WHERE   a.language_cd = @defaultlang_cd ")
            sbSql.Append(" AND     a.program_id  = @program_id ")
            sbSql.Append(" ORDER BY  a.label_seq ")

            objCmd = New SqlCommand(sbSql.ToString, objCon)
            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@language_cd", SqlDbType.Char, 2).Value = strLangCd
                .Parameters.Add("@defaultlang_cd", SqlDbType.Char, 2).Value = CdCst.LanguageCd.DefaultLang
                .Parameters.Add("@program_id", SqlDbType.VarChar, 25).Value = strPgmId
            End With

            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(fncSelectPageAllLabels)

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            'DBオブジェクト破棄
            If Not objAdp Is Nothing Then objAdp.Dispose()
            objAdp = Nothing
            If Not objCmd Is Nothing Then objCmd.Dispose()
            objCmd = Nothing
            sbSql = Nothing
        End Try
    End Function

    ''' <summary>
    ''' ラベルデータの検索
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strPgmId">プログラムＩＤ</param>
    ''' <param name="strLangCd">言語コード </param>
    ''' <param name="strLblDiv"> ラベル区分       L:画面ラベル  B:ボタンラベル</param>
    ''' <param name="intLblSeq">ラベル番号</param>
    ''' <returns></returns>
    ''' <remarks>ラベルマスタより引数.プログラムID、言語コード、ラベル区分、ラベル番号に該当するラベルを取得する</remarks>
    Public Shared Function fncSelectLabelById(ByVal objCon As SqlConnection, ByVal strPgmId As String, _
                                  ByVal strLangCd As String, ByVal strLblDiv As String, ByVal intLblSeq As Integer) As String
        Dim cm As SqlCommand = objCon.CreateCommand
        Dim dr As SqlDataReader = Nothing
        Dim ret As String = ""
        fncSelectLabelById = String.Empty

        Try
            cm.CommandText = " SELECT  ISNULL(b.label_content, a.label_content) AS label_content " & _
                             " FROM    sales.kh_label_mst2 a " & _
                             " LEFT JOIN  sales.kh_label_mst2 b " & _
                             " ON      a.program_id  = b.program_id " & _
                             " AND     a.label_div   = b.label_div " & _
                             " AND     a.label_seq   = b.label_seq " & _
                             " AND     b.language_cd = @language_cd " & _
                             " WHERE   a.language_cd = @defaultlang_cd " & _
                             " AND     a.program_id  = @program_id " & _
                             " AND     a.label_div   = @label_div " & _
                             " AND     a.label_seq   = @label_seq " & _
                             " ORDER BY  a.label_seq "
            With cm.Parameters
                .Add("@language_cd", SqlDbType.Char, 2).Value = strLangCd
                .Add("@defaultlang_cd", SqlDbType.Char, 2).Value = CdCst.LanguageCd.DefaultLang
                .Add("@program_id", SqlDbType.VarChar, 25).Value = strPgmId
                .Add("@label_div", SqlDbType.Char, 1).Value = strLblDiv
                .Add("@label_seq", SqlDbType.Int).Value = intLblSeq
            End With
            dr = cm.ExecuteReader
            While dr.Read
                ret = dr("label_content")
                Exit While
            End While

            Return ret
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If Not dr Is Nothing Then If Not dr.IsClosed Then dr.Close()
        End Try
    End Function

End Class
