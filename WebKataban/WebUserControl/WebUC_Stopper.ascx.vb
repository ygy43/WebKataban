Public Class WebUC_Stopper
    Inherits KHBase

#Region "プロパティ"
    Public Event BacktoYouso()
#End Region

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmInit()
        Me.OnLoad(Nothing)
    End Sub

    ''' <summary>
    ''' ロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Me.Visible Then Exit Sub
        If Me.objUserInfo.UserId Is Nothing Then Exit Sub
        Try
            Call objKtbnStrc.subSelKtbnInfoGet(objCon, Me.objUserInfo.UserId, Me.objLoginInfo.SessionId)
            Me.lblSeriesNm.Text = objKtbnStrc.strcSelection.strGoodsNm

            img1.Visible = False
            img2.Visible = False
            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "LCG", "LCG-Q"
                    img1.Visible = True
                Case "LCR", "LCR-Q"
                    img2.Visible = True
            End Select
            Call SetAllFontName(Me)
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' OKボタンイベント
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOK_Click(sender As Object, e As System.EventArgs) Handles btnOK.Click
        RaiseEvent BacktoYouso()
    End Sub
End Class