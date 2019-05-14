Imports WebKataban.ClsCommon
Imports System.Data.SqlClient

Public Class MenuBLL
    Private Property dllMenu As New MenuDAL

    ''' <summary>
    ''' 通知情報の取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function fncSelectInformation(ByVal objConBase As SqlConnection, ByVal strSelLang As String) As ArrayList
        Dim dt As New DataTable
        Dim blnResult As New ArrayList
        Dim strDispLang As String

        Try
            'ログイン情報の取得
            dt = dllMenu.fncSelectInformation(objConBase, strSelLang)

            '選択言語のメッセージが1件でもあれば、その言語のメッセージを設定し、なければデフォルトのメッセージを設定する
            If dt.Rows.Count > 0 Then
                '言語初期設定
                strDispLang = CdCst.LanguageCd.DefaultLang

                For intLoopCnt = 0 To dt.Rows.Count - 1
                    If dt.Rows(intLoopCnt).Item("language_cd").Trim = strSelLang Then
                        strDispLang = strSelLang
                        Exit For
                    End If
                Next
                For intLoopCnt = 0 To dt.Rows.Count - 1
                    If dt.Rows(intLoopCnt).Item("language_cd").Trim = strDispLang Then
                        'メッセージ設定
                        blnResult.Add(dt.Rows(intLoopCnt).Item("message"))
                    End If
                Next
            End If

        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return blnResult
    End Function
End Class
