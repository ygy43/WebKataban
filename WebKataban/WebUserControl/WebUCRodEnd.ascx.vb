Imports System.Data.SqlClient
Imports WebKataban.ClsCommon

Public Class WebUCRodEnd
    Inherits System.Web.UI.UserControl

#Region "プロパティ"
    '固定値
    Private Const CST_BLANK As String = CdCst.Sign.Blank
    Private Const ID_Label As String = CdCst.Lbl.Name.Label

    'プロパティ設定値
    Private strLangCd As String                             '言語コード
    Private strRodPtn As String                             'ロッド先端パターン
    Private strPtnNo As String                              'パターンNo.
    Private strImageUrl As String                           'イメージURL
    Private strExtFrm() As String                           '外径種類
    Private strDispExtFrm() As String                       '表示外径種類
    Private strNormalVal() As String                        '標準寸法
    Private strActNormalVal() As String                     '実標準寸法
    Private strInputDiv() As String                         '入力区分
    Private strSltVal() As String                           '選択可能寸法
    Private strActSltVal() As String                        '実選択可能寸法
    Private strJsName() As String                           'Javascript名
    Private strEditDiv As String                            '小数点区分
    Private hshtSelSize As Hashtable = Nothing               '特注寸法
    Private strSelOtherVal As String = String.Empty         'その他寸法
    Private bolEnableFlg As Boolean = True                  '入力可能フラグ

    Dim objCon_ As New SqlConnection
#End Region

#Region " Method "

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Page_Init(sender As Object, e As System.EventArgs) Handles Me.Init
        Call SetAttributes(Label4, 0)
        Call SetAttributes(CtlCharText1, 3)
        CtlCharText1.Font.Name = GetFontName(strLangCd)
    End Sub

    ''' <summary>
    ''' 初期処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            '初期画面設定
            Call Me.subSetInitScreen()
            'ラベル設定
            objCon_ = New SqlClient.SqlConnection(My.Settings.connkhdb)
            objCon_.Open()
            Call KHLabelCtl.subSetLabel(objCon_, CdCst.PgmId.KHRodEnd, strLangCd, Me) 'Label取得
        Catch ex As Exception
            'エラー画面に遷移する
            'Call clsPageTrn.subSubmitError(Me.Page, ex)
        Finally
            objCon_.Close()
            objCon_ = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' 初期画面設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub subSetInitScreen()
        Dim objRow As TableRow
        Dim objCell As TableCell
        'Dim objLabel As Label
        Dim objText As TextBox
        Dim objText_input As CtlNumText
        Dim objDrop As DropDownList
        Dim objOption As ListItem
        Dim objImage As Image
        Dim intCellWidth(2) As Integer
        Dim strDropArray() As String
        Dim intLoopCnt1 As Integer
        Dim intLoopCnt2 As Integer
        Dim intLoopCnt3 As Integer

        Dim strParent As String = "ctl00_ContentDetail_WebUC_RodEnd_"

        Try
            '寸法表セル幅
            intCellWidth(0) = 40
            intCellWidth(1) = 90
            intCellWidth(2) = 120

            'その他寸法エリア初期化
            Me.Label4.Visible = False
            Me.CtlCharText1.Visible = False
            'イメージ
            If strImageUrl.Trim.Length <> 0 Then
                objRow = New TableRow
                objCell = New TableCell
                objImage = New Image

                objImage.ImageUrl = strImageUrl
                objCell.Controls.Add(objImage)
                objRow.Controls.Add(objCell)
                Me.tblImg.Rows.Add(objRow)
            End If
            'その他寸法表示
            If strRodPtn.Trim = CdCst.RodEndCstmOrder.OtherSize Then
                Me.Label4.Visible = True
                Me.CtlCharText1.Visible = True
                Me.CtlCharText1.ID = "OtherProd"
                If strSelOtherVal.Trim <> "" Then
                    Me.CtlCharText1.Text = strSelOtherVal.Trim
                End If
                If Not bolEnableFlg Then
                    Me.CtlCharText1.Enabled = False
                End If
                '寸法表
            ElseIf UBound(strExtFrm) <> 0 Then
                'ヘッダー行
                objRow = New TableRow
                For intLoopCnt1 = 1 To 3
                    objCell = New TableCell
                    objText = New TextBox

                    'ID設定
                    objText.ID = ID_Label & intLoopCnt1
                    SetAttributes(objText, 1)
                    objText.Width = intCellWidth(intLoopCnt1 - 1)
                    objText.Font.Name = GetFontName(strLangCd)
                    objCell.Controls.Add(objText)
                    objRow.Controls.Add(objCell)
                Next
                Me.TblLst.Rows.Add(objRow)

                '明細行
                For intLoopCnt1 = 1 To UBound(strExtFrm)
                    '行を追加する
                    objRow = New TableRow

                    '実選択可能KK寸法/標準A寸法設定
                    If strExtFrm(intLoopCnt1) = CdCst.RodEndCstmOrder.FrmKK Then
                        Me.HdnStdKK.Value = strActNormalVal(intLoopCnt1)
                        Me.HdnSltKK.Value = strSltVal(intLoopCnt1)
                        Me.HdnActSltKK.Value = strActSltVal(intLoopCnt1)
                        Me.HdnRowKK.Value = intLoopCnt1
                    ElseIf strExtFrm(intLoopCnt1) = CdCst.RodEndCstmOrder.FrmA Then
                        Me.HdnStdA.Value = strNormalVal(intLoopCnt1)
                        Me.HdnRowA.Value = intLoopCnt1
                    ElseIf strExtFrm(intLoopCnt1) = CdCst.RodEndCstmOrder.FrmC Then
                        Me.HdnStdC.Value = strNormalVal(intLoopCnt1)
                        Me.HdnRowC.Value = intLoopCnt1
                    End If

                    For intLoopCnt2 = 1 To 3
                        objCell = New TableCell

                        Select Case intLoopCnt2
                            Case 1
                                '外径種類
                                objCell.Width = intCellWidth(intLoopCnt2 - 1)

                                'ラベル追加
                                objText = New TextBox
                                SetAttributes(objText, 1)
                                objText.Width = intCellWidth(intLoopCnt2 - 1)
                                objText.Text = strDispExtFrm(intLoopCnt1)
                                objText.Font.Name = GetFontName(strLangCd)
                                objCell.Controls.Add(objText)
                            Case 2
                                '標準寸法
                                objCell.Width = intCellWidth(intLoopCnt2 - 1)

                                'ラベル追加
                                objText = New TextBox
                                SetAttributes(objText, 1)
                                objText.Width = intCellWidth(intLoopCnt2 - 1)
                                objText.Text = strNormalVal(intLoopCnt1)
                                objText.Font.Name = GetFontName(strLangCd)
                                objCell.Controls.Add(objText)
                            Case 3
                                '特注寸法
                                objCell.Width = intCellWidth(intLoopCnt2 - 1)

                                '入力区分によってラベル/テキスト/ドロップダウンを追加
                                Select Case strInputDiv(intLoopCnt1)
                                    Case CdCst.RodEndCstmOrder.Label
                                        objText = New TextBox
                                        SetAttributes(objText, 2)
                                        objText.Width = intCellWidth(intLoopCnt2 - 1)
                                        objText.ID = "Prod" & intLoopCnt1
                                        If hshtSelSize IsNot Nothing Then
                                            If hshtSelSize(strExtFrm(intLoopCnt1)).trim.length <> 0 Then
                                                objText.Text = hshtSelSize(strExtFrm(intLoopCnt1)).trim
                                            Else
                                                objText.Text = strNormalVal(intLoopCnt1)
                                            End If
                                        Else
                                            objText.Text = strNormalVal(intLoopCnt1)
                                        End If
                                        objText.TabIndex = -1
                                        objText.Font.Name = GetFontName(strLangCd)
                                        objCell.Controls.Add(objText)
                                    Case CdCst.RodEndCstmOrder.Text
                                        objText_input = New CtlNumText
                                        objText_input.Width = intCellWidth(intLoopCnt2 - 1)
                                        SetAttributes(objText_input, 3)
                                        objText_input.ID = "Prod" & intLoopCnt1
                                        objText_input.EditDiv = strEditDiv
                                        objText_input.CheckNgProc = "fncDispErrMsg('" & strParent & "');"
                                        objText_input.DecLen = 1
                                        objText_input.AllowMinus = False
                                        objText_input.AllowZero = False
                                        If hshtSelSize IsNot Nothing Then
                                            If hshtSelSize.ContainsKey(strExtFrm(intLoopCnt1)) Then
                                                If hshtSelSize(strExtFrm(intLoopCnt1)).trim.length <> 0 Then
                                                    objText_input.Text = hshtSelSize(strExtFrm(intLoopCnt1)).trim
                                                End If
                                            End If
                                        End If
                                        If Not bolEnableFlg Then
                                            objText_input.Text = CdCst.Sign.Blank
                                            objText_input.ReadOnly = True
                                            objText_input.TabIndex = -1
                                        End If
                                        'javascript追加
                                        If strJsName(intLoopCnt1) <> "" Then
                                            objText_input.CheckOkProc = "f_RodEnd_OnBlur('" & strParent & Me.ID & "','" & strEditDiv & "');"
                                        End If
                                        objText_input.Font.Name = GetFontName(strLangCd)
                                        objCell.Controls.Add(objText_input)
                                    Case CdCst.RodEndCstmOrder.Drop
                                        objDrop = New DropDownList
                                        SetAttributes(objDrop, 0)
                                        objDrop.AutoPostBack = False
                                        objDrop.Width = intCellWidth(intLoopCnt2 - 1) + 5
                                        objDrop.ID = "Prod" & intLoopCnt1
                                        strDropArray = Split(strSltVal(intLoopCnt1), CdCst.Sign.Delimiter.Comma)
                                        For intLoopCnt3 = 0 To strDropArray.Length - 1
                                            objOption = New ListItem
                                            objOption.Text = strDropArray(intLoopCnt3)
                                            objDrop.Items.Add(objOption)
                                        Next
                                        If hshtSelSize IsNot Nothing Then
                                            If hshtSelSize.ContainsKey(strExtFrm(intLoopCnt1)) Then
                                                If hshtSelSize(strExtFrm(intLoopCnt1)).trim.length <> 0 Then
                                                    objDrop.Text = hshtSelSize(strExtFrm(intLoopCnt1)).trim
                                                End If
                                            End If
                                        End If
                                        If Not bolEnableFlg Then
                                            objDrop.Text = CdCst.Sign.Blank
                                            objDrop.Enabled = False
                                        End If
                                        'javascript追加
                                        If strJsName(intLoopCnt1) <> "" Then
                                            objDrop.Attributes.Add("onblur", "f_RodEnd_OnBlur('" & strParent & Me.ID & "','" & strEditDiv & "');")
                                        End If
                                        objDrop.Font.Name = GetFontName(strLangCd)
                                        objCell.Controls.Add(objDrop)
                                End Select
                        End Select
                        objRow.Cells.Add(objCell)
                    Next
                    Me.TblLst.Rows.Add(objRow)
                Next
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

#Region " Property "

    ''' <summary>
    ''' 表示言語の設定・取得の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LangCd() As String
        Get
            Return Me.strLangCd
        End Get
        Set(ByVal value As String)
            Me.strLangCd = value
        End Set
    End Property

    ''' <summary>
    ''' ラジオボタンラベルの設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RodPtn() As String
        Get
            Return Me.strRodPtn
        End Get
        Set(ByVal value As String)
            Me.strRodPtn = value
        End Set
    End Property

    ''' <summary>
    ''' パターンNo.の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property PtnNo() As String
        Get
            Return Me.strPtnNo
        End Get
        Set(ByVal value As String)
            Me.strPtnNo = value
        End Set
    End Property

    ''' <summary>
    ''' イメージパスの設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ImageUrl() As String
        Get
            Return Me.strImageUrl
        End Get
        Set(ByVal value As String)
            Me.strImageUrl = value
        End Set
    End Property

    ''' <summary>
    ''' 外径種類の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ExtFrm() As String()
        Get
            Return Me.strExtFrm
        End Get
        Set(ByVal value As String())
            Me.strExtFrm = value
        End Set
    End Property

    ''' <summary>
    ''' 表示外径種類の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DispExtFrm() As String()
        Get
            Return Me.strDispExtFrm
        End Get
        Set(ByVal value As String())
            Me.strDispExtFrm = value
        End Set
    End Property

    ''' <summary>
    ''' 標準寸法の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property NormalVal() As String()
        Get
            Return Me.strNormalVal
        End Get
        Set(ByVal value As String())
            Me.strNormalVal = value
        End Set
    End Property

    ''' <summary>
    ''' 実標準寸法の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ActNormalVal() As String()
        Get
            Return Me.strActNormalVal
        End Get
        Set(ByVal value As String())
            Me.strActNormalVal = value
        End Set
    End Property

    ''' <summary>
    ''' Javascript名の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property JSName() As String()
        Get
            Return Me.strJsName
        End Get
        Set(ByVal value As String())
            Me.strJsName = value
        End Set
    End Property

    ''' <summary>
    ''' 入力区分の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property InputDiv() As String()
        Get
            Return Me.strInputDiv
        End Get
        Set(ByVal value As String())
            Me.strInputDiv = value
        End Set
    End Property

    ''' <summary>
    ''' 選択可能寸法の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SltVal() As String()
        Get
            Return Me.strSltVal
        End Get
        Set(ByVal value As String())
            Me.strSltVal = value
        End Set
    End Property

    ''' <summary>
    ''' 実選択可能寸法の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ActSltVal() As String()
        Get
            Return Me.strActSltVal
        End Get
        Set(ByVal value As String())
            Me.strActSltVal = value
        End Set
    End Property

    ''' <summary>
    ''' 特注寸法の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SelValInfo() As Hashtable
        Get
            Return Me.hshtSelSize
        End Get
        Set(ByVal value As Hashtable)
            Me.hshtSelSize = value
        End Set
    End Property

    ''' <summary>
    ''' その他寸法の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SelOtherVal() As String
        Get
            Return Me.strSelOtherVal
        End Get
        Set(ByVal value As String)
            Me.strSelOtherVal = value
        End Set
    End Property

    ''' <summary>
    ''' 編集区分の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EditDiv() As String
        Get
            Return strEditDiv
        End Get
        Set(ByVal value As String)
            strEditDiv = value
        End Set
    End Property

    ''' <summary>
    ''' 使用可否の設定・取得
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property EnableFlg() As Boolean
        Get
            Return bolEnableFlg
        End Get
        Set(ByVal value As Boolean)
            Me.bolEnableFlg = value
        End Set
    End Property

    ''' <summary>
    ''' メッセージの表示・非表示
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property VisibleMsg1() As Boolean
        Get
            Return Me.Label5.Visible
        End Get
        Set(ByVal value As Boolean)
            Me.Label5.Visible = value
        End Set
    End Property

    ''' <summary>
    ''' メッセージの表示・非表示
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property VisibleMsg2() As Boolean
        Get
            Return Me.Label6.Visible
        End Get
        Set(ByVal value As Boolean)
            Me.Label6.Visible = value
        End Set
    End Property

#End Region

End Class