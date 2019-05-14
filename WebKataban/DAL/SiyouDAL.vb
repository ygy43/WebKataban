Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class SiyouDAL
    ''' <summary>
    ''' 品名マスタの取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strLanguage"></param>
    ''' <param name="strSpecNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LoadPositionData(ByVal objCon As SqlConnection, strLanguage As String, _
                                      strSpecNo As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL作成
            sbSql.Append("SELECT EN.label_content   AS df_label_content, ")
            sbSql.Append("		 EN.item_num        AS df_item_num,      ")
            sbSql.Append("		 EN.item_div        AS df_item_div,      ")
            sbSql.Append("		 EN.others          AS df_others,        ")
            sbSql.Append("		 IT.label_content,                       ")
            sbSql.Append("		 IT.item_num,                            ")
            sbSql.Append("		 IT.item_div,                            ")
            sbSql.Append("		 IT.others,                              ")
            sbSql.Append("		 IT.label_seq                            ")
            sbSql.Append("FROM	 kh_item_mst EN                          ")
            sbSql.Append("LEFT OUTER JOIN kh_item_mst IT                 ")
            sbSql.Append("	ON	 IT.language_cd	= @LangCd                ")
            sbSql.Append("	AND	 EN.spec_no		= IT.spec_no             ")
            sbSql.Append("	AND	 EN.label_seq	= IT.label_seq           ")
            sbSql.Append("WHERE	 EN.language_cd	= @DefaultLang           ")
            sbSql.Append("AND	 EN.spec_no	    = @SpecNo                ")
            sbSql.Append("ORDER BY EN.label_seq                          ")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@DefaultLang", SqlDbType.Char, 2).Value = CdCst.LanguageCd.DefaultLang
                .Parameters.Add("@LangCd", SqlDbType.Char, 2).Value = strLanguage
                .Parameters.Add("@SpecNo", SqlDbType.Char, 2).Value = strSpecNo
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

    ''' <summary>
    ''' マニホールド画面仕様を取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSpecNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LoadComboData(ByVal objCon As SqlConnection, strSpecNo As String) As DataTable
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL作成
            sbSql.Append("SELECT * ")
            sbSql.Append("FROM dbo.MonifoldKataban ")
            sbSql.Append("WHERE (SpecNo = @SpecNo) AND (Delete_Flg IS NULL OR Delete_Flg <> '9') ")
            sbSql.Append("ORDER BY KeyID")

            'DB接続文字列の取得
            objCmd = New SqlCommand(sbSql.ToString, objCon)

            With objCmd
                .CommandType = CommandType.Text
                .Parameters.Add("@SpecNo", SqlDbType.Char, 2).Value = strSpecNo
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

    ''' <summary>
    ''' 品名マスタ取得
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strSpecNo"></param>
    ''' <param name="dtSpecItem"></param>
    ''' <param name="dtContent"></param>
    ''' <remarks></remarks>
    Public Shared Sub subSQL_ItemMst(objCon As SqlConnection, ByVal strSpecNo As String, _
                                     ByRef dtSpecItem As DataTable, ByRef dtContent As DataTable, _
                                     Optional ByVal strLang As String = "ja")
        Dim objAdp As SqlDataAdapter = Nothing
        Dim sbSql As New StringBuilder
        Dim objCmd As SqlCommand
        Dim dtResult As New DataTable

        Try
            'SQL作成
            sbSql.Append("SELECT EN.label_content   AS df_label_content, ")
            sbSql.Append("		 EN.item_num        AS df_item_num, ")
            sbSql.Append("		 EN.item_div        AS df_item_div, ")
            sbSql.Append("		 EN.others          AS df_others, ")
            sbSql.Append("		 IT.label_content, ")
            sbSql.Append("		 IT.item_num, ")
            sbSql.Append("		 IT.item_div, ")
            sbSql.Append("		 IT.others ")
            sbSql.Append("FROM	 kh_item_mst EN ")
            sbSql.Append("LEFT OUTER JOIN kh_item_mst IT ")
            sbSql.Append("	ON	 IT.language_cd	= @LangCd ")
            sbSql.Append("	AND	 EN.spec_no		= IT.spec_no ")
            sbSql.Append("	AND	 EN.label_seq	= IT.label_seq ")
            sbSql.Append("WHERE	 EN.language_cd	= @DefaultLang ")
            sbSql.Append("AND	 EN.spec_no	    = @SpecNo ")
            sbSql.Append("ORDER BY EN.label_seq ")

            objCmd = objCon.CreateCommand

            With objCmd
                .CommandText = sbSql.ToString

                .Parameters.Add("@DefaultLang", SqlDbType.Char, 2).Value = CdCst.LanguageCd.DefaultLang
                .Parameters.Add("@LangCd", SqlDbType.Char, 2).Value = strLang
                .Parameters.Add("@SpecNo", SqlDbType.Char, 2).Value = strSpecNo
            End With
            '実行
            objAdp = New SqlDataAdapter(objCmd)
            objAdp.Fill(dtResult)

            For Each objRdr As DataRow In dtResult.Rows
                Dim strProdNm As String = String.Empty
                Dim intItemCnt As Integer
                Dim strItemDiv As String = String.Empty
                Dim rowSpecItem As DataRow
                Dim rowContent As DataRow

                If IsDBNull(objRdr("label_content")) Then
                    strProdNm = objRdr("df_label_content").ToString
                    intItemCnt = CInt(objRdr("df_item_num").ToString)
                    strItemDiv = objRdr("df_item_div").ToString
                Else
                    strProdNm = objRdr("label_content").ToString
                    intItemCnt = CInt(objRdr("item_num").ToString)
                    strItemDiv = objRdr("item_div").ToString
                End If

                '項目区分(=1：記号　=2：形番)によって保持ﾃｰﾌﾞﾙを分ける
                Select Case strItemDiv
                    Case "1"
                        rowSpecItem = dtSpecItem.NewRow
                        rowSpecItem(CdCst.TblSpecItem.ProdNm) = strProdNm
                        rowSpecItem(CdCst.TblSpecItem.ItemCnt) = intItemCnt
                        rowSpecItem(CdCst.TblSpecItem.ItemDiv) = strItemDiv
                        dtSpecItem.Rows.Add(rowSpecItem)
                    Case "2"
                        rowContent = dtContent.NewRow
                        rowContent(CdCst.TblSpecItem.ProdNm) = strProdNm
                        rowContent(CdCst.TblSpecItem.ItemCnt) = intItemCnt
                        rowContent(CdCst.TblSpecItem.ItemDiv) = strItemDiv
                        dtContent.Rows.Add(rowContent)
                    Case Else
                End Select
            Next

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            If objAdp IsNot Nothing Then
                objAdp.Dispose()
                objAdp = Nothing
            End If
            objCmd = Nothing
        End Try
    End Sub
End Class
