'************************************************************************************
'*  ProgramID  �FKHPriceR5
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2011/10/21   �쐬�ҁFY.Tachi
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �F�Ȕz���u���b�N�}�j�z�[���h             �l�m�R�p�O�^�l�s�R�p�O
'*
'************************************************************************************
Module KHPriceR5
    'RM0911XXX 2009/11/16 Y.Miura �_�~�[�u���b�N�ǉ�
    Private Structure ItemNum
        Private dummy As Integer
        Public Const Elect1 As Integer = 1
        Public Const Elect2 As Integer = 2
        Public Const Wiring As Integer = 3
        Public Const Valve1 As Integer = 4
        Public Const Valve2 As Integer = 5
        Public Const Valve3 As Integer = 6
        Public Const Valve4 As Integer = 7
        Public Const Valve5 As Integer = 8
        Public Const Valve6 As Integer = 9
        Public Const Valve7 As Integer = 10
        Public Const Dummy1 As Integer = 11
        Public Const Dummy2 As Integer = 12
        Public Const Exhaust1 As Integer = 13
        Public Const Exhaust2 As Integer = 14
        Public Const Exhaust3 As Integer = 15
        Public Const Exhaust4 As Integer = 16
        Public Const Regulat1 As Integer = 17
        Public Const Regulat2 As Integer = 18
        Public Const EndL As Integer = 19
        Public Const EndR As Integer = 20
        Public Const Plug1 As Integer = 21
        Public Const Plug2 As Integer = 22
        Public Const Plug3 As Integer = 23
        Public Const Plug4 As Integer = 24
        Public Const Inspect1 As Integer = 25
        Public Const Inspect2 As Integer = 26
        Public Const Inspect3 As Integer = 27
        Public Const Inspect4 As Integer = 28

    End Structure

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal, _
                                   Optional ByRef strPriceDiv() As String = Nothing)

        Dim strOpArray() As String
        Dim intLoopCnt As Integer

        Dim intStationQty As Integer = 0
        Dim intValveQty As Integer = 0

        Dim intValve3PQty As Integer = 0
        Dim intValve3PDualQty As Integer = 0
        Dim intValve4PQty As Integer = 0

        'RM0911XXX 2009/11/16 Y.Miura �_�~�[�u���b�N�ǉ�
        Dim intValve307Qty As Integer = 0           '�o���u�u���b�N 7mm�̐�
        Dim intValve310Qty As Integer = 0           '�o���u�u���b�N10mm�̐�
        Dim intValve407Qty As Integer = 0           '�o���u�u���b�N 7mm�̐�
        Dim intValve410Qty As Integer = 0           '�o���u�u���b�N10mm�̐�

        Dim strOptionA As String = String.Empty     '�I�v�V����A

        Dim bolOptionS As Boolean = False
        Dim bolOptionSA As Boolean = False
        Dim bolOptionC As Boolean = False

        ' 2008/12/03 �ǉ�
        Dim ItemKiriIchikbn As String = String.Empty        '�؊��ʒu�敪
        Dim ItemSosakbn As String = String.Empty            '����敪
        Dim ItemKokei As String = String.Empty              '�ڑ����a
        Dim ItemSyudoSochi As String = String.Empty         '�蓮���u
        Dim ItemHaisen As String = String.Empty             '�z���ڑ�
        Dim ItemOption As String = String.Empty             '�I�v�V����
        Dim ItemRensu As String = String.Empty              '�A��
        Dim ItemDenatsu As String = String.Empty            '�d��

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)
            ReDim strPriceDiv(0)

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "MN3Q0", "MT3Q0"
                    ItemKiriIchikbn = objKtbnStrc.strcSelection.strOpSymbol(1).Trim '�؊��ʒu�敪
                    ItemSosakbn = objKtbnStrc.strcSelection.strOpSymbol(2).Trim     '����敪
                    ItemKokei = objKtbnStrc.strcSelection.strOpSymbol(3).Trim       '�ڑ����a
                    ItemSyudoSochi = objKtbnStrc.strcSelection.strOpSymbol(4).Trim  '�蓮���u
                    ItemHaisen = objKtbnStrc.strcSelection.strOpSymbol(5).Trim      '�z���ڑ�
                    ItemOption = objKtbnStrc.strcSelection.strOpSymbol(6).Trim      '�I�v�V����
                    ItemRensu = objKtbnStrc.strcSelection.strOpSymbol(7).Trim       '�A��
                    ItemDenatsu = objKtbnStrc.strcSelection.strOpSymbol(8).Trim     '�d��
            End Select

            '�A���ݒ�
            intStationQty = CInt(ItemRensu)

            'RM0911XXX 2009/11/16 Y.Miura �_�~�[�u���b�N�ǉ�
            '�I�v�V�������Z���i�L�[
            strOpArray = Split(ItemOption, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case ""
                    Case "A"
                        strOptionA = strOpArray(intLoopCnt).Trim
                    Case "E"
                    Case Else
                End Select
            Next

            For intLoopCnt = 1 To objKtbnStrc.strcSelection.strOptionKataban.Length - 1
                '�`�ԁE�g�p�������݂���ꍇ
                If objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim <> "" And _
                   objKtbnStrc.strcSelection.intQuantity(intLoopCnt) > 0 Then
                    bolOptionS = False
                    bolOptionSA = False
                    bolOptionC = False

                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                        Case CdCst.Manifold.InspReportJp.Japanese, _
                             CdCst.Manifold.InspReportJp.English, _
                             CdCst.Manifold.InspReportEn.Japanese, _
                             CdCst.Manifold.InspReportEn.English
                        Case Else
                            Select Case intLoopCnt
                                'Case 1 To 2
                                Case ItemNum.Elect1 To ItemNum.Elect2
                                    '�d���u���b�N
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                    'Case 3
                                Case ItemNum.Wiring
                                    '�ʔz��

                                    'Case 4 To 11
                                Case ItemNum.Valve1 To ItemNum.Valve7
                                    '�o���u�u���b�N
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    '�o���u�����J�E���g
                                    intValveQty = intValveQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    '�o���u�����J�E���g
                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 2, 1)
                                        Case "3"
                                            Select Case Left(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5)
                                                Case "MN3Q0", "MT3Q0"
                                                    Select Case Mid(objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim, 5, 2)
                                                        Case "66"
                                                            intValve3PDualQty = intValve3PDualQty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                                    End Select
                                            End Select
                                    End Select

                                    'RM0911XXX 2009/11/16 Y.Miura �_�~�[�u���b�N�ǉ�
                                    Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Substring(0, 5)
                                        Case "MN3Q0", "MT3Q0"
                                            Select Case objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim.Substring(1, 1)
                                                Case "3"
                                                    intValve310Qty = intValve310Qty + objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                                            End Select
                                    End Select
                                    'RM0911XXX 2009/11/16 Y.Miura �_�~�[�u���b�N�ǉ�
                                Case ItemNum.Dummy1 To ItemNum.Dummy2
                                    '�_�~�[�u���b�N
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'Case 13 To 16
                                Case ItemNum.Exhaust1 To ItemNum.Exhaust4
                                    '���r�C�u���b�N
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'Case 17 To 18
                                Case ItemNum.Regulat1 To ItemNum.Regulat2
                                    '���M�����[�^�u���b�N
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'Case 19 To 20
                                Case ItemNum.EndL To ItemNum.EndR
                                    '�G���h�u���b�N
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'Case 21 To 24
                                Case ItemNum.Plug1 To ItemNum.Plug4
                                    '�u�����N�v���O���T�C�����T
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = "N3Q0-" & objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)

                                    'Case 26 To 28
                                Case ItemNum.Inspect2 To ItemNum.Inspect4
                                    '�������я����P�[�u�����p�聕�R�l�N�^���\�P�b�g
                                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strOptionKataban(intLoopCnt).Trim
                                    decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.intQuantity(intLoopCnt)
                            End Select
                    End Select
                End If
            Next

            'MN3Q0�V���[�Y�̂�
            If InStr(1, objKtbnStrc.strcSelection.strSeriesKataban.Trim, "N") <> 0 Then
                '��t���[���������Z���i�L�[
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "N4E0-BAA"
                decOpAmount(UBound(decOpAmount)) = objKtbnStrc.strcSelection.decDinRailLength
                strPriceDiv(UBound(strPriceDiv)) = CdCst.PriceAccDiv.DINRail
            End If

            '�_�C���N�g�}�E���g�������Z���i�L�[
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "MT3Q0"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = "MT3Q0-DM"
                    decOpAmount(UBound(decOpAmount)) = 1
            End Select

            '�m�����b�N���蓮���u���Z���i�L�[
            If ItemSyudoSochi = "M" Then
                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                strOpRefKataban(UBound(strOpRefKataban)) = "N3Q0" & CdCst.Sign.Hypen & ItemSyudoSochi
                decOpAmount(UBound(decOpAmount)) = intValveQty
            End If

            '�I�v�V�������Z���i�L�[
            strOpArray = Split(ItemOption, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "F"
                        '�� "N3Q0-F"
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        ReDim Preserve strPriceDiv(UBound(strPriceDiv) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = "N3Q0" & CdCst.Sign.Hypen & _
                                                                   strOpArray(intLoopCnt)
                        decOpAmount(UBound(decOpAmount)) = intValveQty
                    Case "P", "N"
                End Select
            Next

        Catch ex As Exception

            Throw ex

        End Try

    End Sub

End Module
