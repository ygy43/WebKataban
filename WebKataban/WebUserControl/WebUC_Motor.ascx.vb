Public Class WebUC_Motor
    Inherits KHBase

#Region "プロパティ"
    Private strLanguage As String    '選択した言語
    Private strSeries As String      '機種
    Private strSeriesName As String  '機種名称
    Public Event BacktoYouso()       '仕様画面へイベント
#End Region

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub frmInit()
        Call subSetAllImageUnvisible()
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
            Me.lblSeriesName.Text = objKtbnStrc.strcSelection.strGoodsNm

            Select Case objKtbnStrc.strcSelection.strSeriesKataban
                Case "ETS", "EBS", "EBR"    'RM1803042_EBS,EBR追加
                    Select Case objKtbnStrc.strcSelection.strKeyKataban
                        Case "A"
                            Select Case objLoginInfo.SelectLang
                                Case "ja"
                                    ImageJA.Visible = True
                                    ImageJA.ImageUrl = "../KHImage/ETS-AMotorJapan.gif"
                                Case "en"
                                    ImageEN.Visible = True
                                    ImageEN.ImageUrl = "../KHImage/ETS-AMotorEnglish.gif"
                                Case "ko"
                                    ImageKO.Visible = True
                                    ImageKO.ImageUrl = "../KHImage/ETS-AMotorKorea.gif"
                                Case "tw"
                                    ImageTW.Visible = True
                                    ImageTW.ImageUrl = "../KHImage/ETS-AMotorTaiwan.gif"
                                Case "zh"
                                    ImageZH.Visible = True
                                    ImageZH.ImageUrl = "../KHImage/ETS-AMotorChina.gif"
                            End Select
                        Case "B"
                            Select Case objLoginInfo.SelectLang
                                Case "ja"
                                    ImageJA.Visible = True
                                    ImageJA.ImageUrl = "../KHImage/ETS-GMotorJapan.gif"
                                Case "en"
                                    ImageEN.Visible = True
                                    ImageEN.ImageUrl = "../KHImage/ETS-GMotorEnglish.gif"
                                Case "ko"
                                    ImageKO.Visible = True
                                    ImageKO.ImageUrl = "../KHImage/ETS-GMotorKorea.gif"
                                Case "tw"
                                    ImageTW.Visible = True
                                    ImageTW.ImageUrl = "../KHImage/ETS-GMotorTaiwan.gif"
                                Case "zh"
                                    ImageZH.Visible = True
                                    ImageZH.ImageUrl = "../KHImage/ETS-GMotorChina.gif"
                            End Select
                        Case "C"
                            Select Case objLoginInfo.SelectLang
                                Case "ja"
                                    ImageJA.Visible = True
                                    ImageJA.ImageUrl = "../KHImage/ETS-PMotorJapan.gif"
                                Case "en"
                                    ImageEN.Visible = True
                                    ImageEN.ImageUrl = "../KHImage/ETS-PMotorEnglish.gif"
                                Case "ko"
                                    ImageKO.Visible = True
                                    ImageKO.ImageUrl = "../KHImage/ETS-PMotorKorea.gif"
                                Case "tw"
                                    ImageTW.Visible = True
                                    ImageTW.ImageUrl = "../KHImage/ETS-PMotorTaiwan.gif"
                                Case "zh"
                                    ImageZH.Visible = True
                                    ImageZH.ImageUrl = "../KHImage/ETS-PMotorChina.gif"
                            End Select
                        Case "D"
                            Select Case objLoginInfo.SelectLang
                                Case "ja"
                                    ImageJA.Visible = True
                                    ImageJA.ImageUrl = "../KHImage/ETS-FMotorJapan.gif"
                                Case "en"
                                    ImageEN.Visible = True
                                    ImageEN.ImageUrl = "../KHImage/ETS-FMotorEnglish.gif"
                                Case "ko"
                                    ImageKO.Visible = True
                                    ImageKO.ImageUrl = "../KHImage/ETS-FMotorKorea.gif"
                                Case "tw"
                                    ImageTW.Visible = True
                                    ImageTW.ImageUrl = "../KHImage/ETS-FMotorTaiwan.gif"
                                Case "zh"
                                    ImageZH.Visible = True
                                    ImageZH.ImageUrl = "../KHImage/ETS-FMotorChina.gif"
                            End Select
                        Case Else
                            Select Case objLoginInfo.SelectLang
                                Case "ja"
                                    ImageJA.Visible = True
                                    ImageJA.ImageUrl = "../KHImage/ETSMotorJapan.gif"
                                Case "en"
                                    ImageEN.Visible = True
                                    ImageEN.ImageUrl = "../KHImage/ETSMotorEnglish.gif"
                                Case "ko"
                                    ImageKO.Visible = True
                                    ImageKO.ImageUrl = "../KHImage/ETSMotorKorea.gif"
                                Case "tw"
                                    ImageTW.Visible = True
                                    ImageTW.ImageUrl = "../KHImage/ETSMotorTaiwan.gif"
                                Case "zh"
                                    ImageZH.Visible = True
                                    ImageZH.ImageUrl = "../KHImage/ETSMotorChina.gif"
                            End Select
                    End Select
                Case "ETV", "ECS", "ECV"
                    Select Case objLoginInfo.SelectLang
                        Case "ja"
                            ImageJA.Visible = True
                            ImageJA.ImageUrl = "../KHImage/ETSMotorJapan.gif"
                        Case "en"
                            ImageEN.Visible = True
                            ImageEN.ImageUrl = "../KHImage/ETSMotorEnglish.gif"
                        Case "ko"
                            ImageKO.Visible = True
                            ImageKO.ImageUrl = "../KHImage/ETSMotorKorea.gif"
                        Case "tw"
                            ImageTW.Visible = True
                            ImageTW.ImageUrl = "../KHImage/ETSMotorTaiwan.gif"
                        Case "zh"
                            ImageZH.Visible = True
                            ImageZH.ImageUrl = "../KHImage/ETSMotorChina.gif"
                    End Select
                Case "ESM"
                    Select Case objLoginInfo.SelectLang
                        Case "ja"
                            ImageJA.Visible = True
                            ImageJA.ImageUrl = "../KHImage/ESMMotorJapan.gif"
                        Case "en"
                            ImageEN.Visible = True
                            ImageEN.ImageUrl = "../KHImage/ESMMotorEnglish.gif"
                        Case "ko"
                            ImageKO.Visible = True
                            ImageKO.ImageUrl = "../KHImage/ESMMotorKorea.gif"
                        Case "tw"
                            ImageTW.Visible = True
                            ImageTW.ImageUrl = "../KHImage/ESMMotorTaiwan.gif"
                        Case "zh"
                            ImageZH.Visible = True
                            ImageZH.ImageUrl = "../KHImage/ESMMotorChina.gif"
                    End Select
                Case "IAVB"
                    ImageIAVB.Visible = True
                    ImageIAVB.ImageUrl = "../KHImage/IAVBPortPosition.gif"
                    'RM1804032_画像表示追加
                Case "EKS"
                    If Me.Session("LabelClick7") = True Then
                        Select Case objLoginInfo.SelectLang
                            Case "ja"
                                ImageJA.Visible = True
                                ImageJA.ImageUrl = "../KHImage/EKSListJapan.gif"
                            Case "en"
                                ImageEN.Visible = True
                                ImageEN.ImageUrl = "../KHImage/EKSListEnglish.gif"
                            Case "ko"
                                ImageKO.Visible = True
                                ImageKO.ImageUrl = "../KHImage/EKSListKorea.gif"
                            Case "tw"
                                ImageTW.Visible = True
                                ImageTW.ImageUrl = "../KHImage/EKSListTaiwan.gif"
                            Case "zh"
                                ImageZH.Visible = True
                                ImageZH.ImageUrl = "../KHImage/EKSListChina.gif"
                        End Select
                    Else
                        Select Case objLoginInfo.SelectLang
                            Case "ja"
                                ImageJA.Visible = True
                                ImageJA.ImageUrl = "../KHImage/ETSMotorJapan.gif"
                            Case "en"
                                ImageEN.Visible = True
                                ImageEN.ImageUrl = "../KHImage/ETSMotorEnglish.gif"
                            Case "ko"
                                ImageKO.Visible = True
                                ImageKO.ImageUrl = "../KHImage/ETSMotorKorea.gif"
                            Case "tw"
                                ImageTW.Visible = True
                                ImageTW.ImageUrl = "../KHImage/ETSMotorTaiwan.gif"
                            Case "zh"
                                ImageZH.Visible = True
                                ImageZH.ImageUrl = "../KHImage/ETSMotorChina.gif"
                        End Select
                    End If
                    Me.Session("LabelClick7") = False
            End Select
            Call SetAllFontName(Me)
        Catch ex As Exception
            AlertMessage(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 要素画面に戻る
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnOK_Click(sender As Object, e As System.EventArgs) Handles btnOK.Click
        RaiseEvent BacktoYouso()
    End Sub

    ''' <summary>
    ''' 全てのイメージを初期化
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetAllImageUnvisible()
        ImageJA.Visible = False
        ImageEN.Visible = False
        ImageKO.Visible = False
        ImageTW.Visible = False
        ImageZH.Visible = False
        ImageIAVB.Visible = False 'RM1610026
    End Sub
End Class