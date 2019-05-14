Public Class Master_Main
    Inherits System.Web.UI.MasterPage

#Region "イベントハンドラ"

    Public Property ReleaseDate As String
    ''' <summary>
    ''' 初期化時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    End Sub

    ''' <summary>
    ''' ロード時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ReleaseDate = Now.Date.ToString("yyyyMMdd")
    End Sub

#End Region

End Class