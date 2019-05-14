Imports System.Data.SqlClient
Imports WebKataban.ClsCommon

Public Class KHLabelCtl
    Private Property bllLabel As New LabelBLL

    ''' <summary>
    ''' 画面ラベル設定
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strFrmID">画面ID</param>
    ''' <param name="strLanguage">選択言語</param>
    ''' <param name="objFrm">画面オブジェクト</param>
    ''' <remarks></remarks>
    Public Shared Sub subSetLabel(ByRef objCon As SqlConnection, strFrmID As String, strLanguage As String, ByRef objFrm As Object)
        Dim dtLblTbl As New DataTable                               'ラベル内容の配列
        Try
            'ラベル設定
            dtLblTbl = fncGetPageAllLabels(objCon, strFrmID, strLanguage)

            Dim dr() As DataRow
            If Not dtLblTbl Is Nothing Then
                dr = dtLblTbl.Select("label_div='" & CdCst.Lbl.Division.Label & "'")  'Label
                For inti As Integer = 0 To dr.Length - 1
                    If Not objFrm.FindControl(CdCst.Lbl.Name.Label & dr(inti)("label_seq").ToString) Is Nothing Then
                        objFrm.FindControl(CdCst.Lbl.Name.Label & dr(inti)("label_seq").ToString).Text = dr(inti)("label_content").ToString
                        SetFontName(objFrm.FindControl(CdCst.Lbl.Name.Label & dr(inti)("label_seq").ToString), strLanguage)

                        SetFontBold(objFrm.FindControl(CdCst.Lbl.Name.Label & dr(inti)("label_seq").ToString))
                    End If
                Next

                dr = dtLblTbl.Select("label_div='" & CdCst.Lbl.Division.Button & "'")  'Button
                For inti As Integer = 0 To dr.Length - 1
                    If Not objFrm.FindControl(CdCst.Lbl.Name.Button & dr(inti)("label_seq").ToString) Is Nothing Then
                        objFrm.FindControl(CdCst.Lbl.Name.Button & dr(inti)("label_seq").ToString).Text = dr(inti)("label_content").ToString

                        SetFontName(objFrm.FindControl(CdCst.Lbl.Name.Button & dr(inti)("label_seq").ToString), strLanguage)

                        SetFontBold(objFrm.FindControl(CdCst.Lbl.Name.Button & dr(inti)("label_seq").ToString))

                        SetButtonJavascript(objFrm.FindControl(CdCst.Lbl.Name.Button & dr(inti)("label_seq").ToString))
                    End If
                Next

                dr = dtLblTbl.Select("label_div='" & CdCst.Lbl.Division.Radio & "'")  'Radio
                If Not objFrm.findcontrol(CdCst.Lbl.Name.RadioButtonList & "1") Is Nothing Then
                    'RadioButtonList
                    Dim radioList As RadioButtonList = CType(objFrm.findcontrol(CdCst.Lbl.Name.RadioButtonList & "1"), RadioButtonList)
                    'For inti As Integer = 0 To dr.Length - 1
                    '    If inti <= radioList.Items.Count Then
                    '        radioList.Items(inti).Text = dr(inti)("label_content").ToString
                    '        radioList.Items(inti).Value = inti
                    '    End If
                    'Next

                    '並び順を　機種、フル形番、仕入品、全て　にする
                    For inti As Integer = 0 To dr.Length - 1
                        If inti <= radioList.Items.Count Then
                            Select Case inti
                                Case 0
                                    radioList.Items(0).Text = dr(0)("label_content").ToString
                                    radioList.Items(0).Value = 0
                                Case 1
                                    radioList.Items(1).Text = dr(1)("label_content").ToString
                                    radioList.Items(1).Value = 1
                                Case 2
                                    radioList.Items(2).Text = dr(3)("label_content").ToString
                                    radioList.Items(2).Value = 3
                                Case 3
                                    radioList.Items(3).Text = dr(2)("label_content").ToString
                                    radioList.Items(3).Value = 2
                            End Select
                        End If
                    Next

                    SetFontName(radioList, strLanguage)
                    SetFontBold(radioList)
                Else
                    'RadioButton
                    For inti As Integer = 0 To dr.Length - 1
                        If Not objFrm.FindControl(CdCst.Lbl.Name.RadioButton & dr(inti)("label_seq").ToString) Is Nothing Then
                            objFrm.FindControl(CdCst.Lbl.Name.RadioButton & dr(inti)("label_seq").ToString).Text = dr(inti)("label_content").ToString
                            SetFontName(objFrm.FindControl(CdCst.Lbl.Name.RadioButton & dr(inti)("label_seq").ToString), strLanguage)

                            SetFontBold(objFrm.FindControl(CdCst.Lbl.Name.RadioButton & dr(inti)("label_seq").ToString))
                        End If
                    Next
                End If

                dr = dtLblTbl.Select("label_div='" & CdCst.Lbl.Division.Title & "'")  'Title
                For inti As Integer = 0 To dr.Length - 1
                    If Not objFrm.FindControl(CdCst.Lbl.Name.Title & dr(inti)("label_seq").ToString) Is Nothing Then
                        objFrm.FindControl(CdCst.Lbl.Name.Title & dr(inti)("label_seq").ToString).Text = dr(inti)("label_content").ToString
                        SetFontName(objFrm.FindControl(CdCst.Lbl.Name.Title & dr(inti)("label_seq").ToString), strLanguage)

                        SetFontBold(objFrm.FindControl(CdCst.Lbl.Name.Title & dr(inti)("label_seq").ToString))
                    End If
                Next
            End If
            dr = Nothing
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        Finally
            dtLblTbl = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' ラベルデータ取り込み
    ''' </summary>
    ''' <param name="objCon"></param>
    ''' <param name="strPgmId">プログラムＩＤ</param>
    ''' <param name="strLangCd">言語コード </param>
    ''' <returns></returns>
    ''' <remarks>
    ''' 引数で渡されたプログラムＩＤ、言語コード、ラベル区分にてラベルマスタを絞り込み
    ''' ラベル内容の配列を構築する。
    ''' </remarks>
    Public Shared Function fncGetPageAllLabels(ByVal objCon As SqlConnection, ByVal strPgmId As String, _
                                 ByVal strLangCd As String) As DataTable
        fncGetPageAllLabels = New DataTable
        Try
            fncGetPageAllLabels = LabelBLL.fncSelectPageAllLabels(objCon, strLangCd, strPgmId)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
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
        Dim strResult As String = String.Empty

        Try
            strResult = LabelBLL.fncSelectLabelById(objCon, strPgmId, strLangCd, strLblDiv, intLblSeq)
        Catch ex As Exception
            WriteErrorLog("E001", ex)
        End Try

        Return strResult
    End Function

    ''' <summary>
    ''' ボタンのonfocusとblurイベント
    ''' </summary>
    ''' <param name="btn"></param>
    ''' <remarks></remarks>
    Private Shared Sub SetButtonJavascript(ByRef btn As Object)
        Dim strJSFocus As String = String.Empty
        Dim strJSBlur As String = String.Empty

        'strJSFocus = strJSFocus & "this.style.color = '#ffcc33'; "
        'strJSBlur = strJSBlur & "this.style.color = '#ffffff'; "

        strJSFocus = strJSFocus & "this.style.color = 'white'; "
        strJSBlur = strJSBlur & "this.style.color = 'black'; "

        CType(btn, Button).Attributes.Add("onfocus", strJSFocus)
        CType(btn, Button).Attributes.Add("onblur", strJSBlur)
    End Sub
End Class
