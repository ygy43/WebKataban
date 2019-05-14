'************************************************************************************
'*  ProgramID  �FKHCylinderSSD2Check
'*  Program��  �F�V�����_�r�r�c�Q�V���[�Y�`�F�b�N���W���[��
'*
'*                                      �쐬���F2008/01/11   �쐬�ҁFNII A.Takahashi
'*
'*  �T�v       �F�r�r�c�Q�^�r�r�c�Q�[�j
'*  �E��tNo�FRM0906034  �񎟓d�r�Ή��@��Ή�
'*                                      �X�V���F2009/08/05   �X�V�ҁFY.Miura
'************************************************************************************
Module KHCylinderSSD2Check

    '********************************************************************************************
    '*�y�֐����z
    '*  fncCheckSelectOption
    '*�y�����z
    '*  �V�����_�`�F�b�N
    '*�y�T�v�z
    '*  �V�����_�r�r�c�Q�V���[�Y���`�F�b�N����
    '*�y�����z
    '*  <Object>       objKtbnStrc          �����`�ԏ��
    '*  <Integer>      intKtbnStrcSeqNo     �`�ԍ\������
    '*  <String>       strOptionSymbol      �I�v�V�����L��
    '*  <String>       strMessageCd         ���b�Z�[�W�R�[�h
    '*�y�߂�l�z
    '*  <Boolean>
    '*�y�C�������z
    '*                                      �X�V���F2008/05/07   �X�V�ҁFT.Sato
    '*  �E��tNo�FRM0802088�Ή��@�o���G�[�V�����i'�c','�l','�p','�w','�x'�j�ǉ��ɔ����C��
    '********************************************************************************************
    Public Function fncCheckSelectOption(ByVal objKtbnStrc As KHKtbnStrc, _
                                         ByRef intKtbnStrcSeqNo As Integer, _
                                         ByRef strOptionSymbol As String, _
                                         ByRef strMessageCd As String) As Boolean

        Try

            fncCheckSelectOption = True

            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim
                Case "SSD2"
                    '��{�x�[�X���Ƀ`�F�b�N
                    Select Case objKtbnStrc.strcSelection.strKeyKataban.Trim
                        '2010/10/05 MOD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
                        Case ""
                            '��{�x�[�X�`�F�b�N
                            If fncStandardBaseCheck(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case "4"
                            'RM0906034 2009/08/05 Y.Miura�@�񎟓d�r�Ή��@��ǉ�
                            'Case "", "4"
                            ''Case ""
                            '2010/10/05 MOD RM1010017(11��VerUP:SSD2�V���[�Y) <---END
                            '�Q���d�r�`�F�b�N
                            If fncP4BaseCheck(objKtbnStrc, _
                                              intKtbnStrcSeqNo, _
                                              strOptionSymbol, _
                                              strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                        Case "D", "E"
                            '�����b�h�x�[�X�`�F�b�N
                            If fncDoubleRodBaseCheck(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If

                            '2010/11/02 MOD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                        Case "K"
                            '���׏d�x�[�X�`�F�b�N
                            If fncHighLoadBaseCheck(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If

                        Case "L"
                            '���׏d�x�[�X�`�F�b�N
                            If fncHighLoadBaseP4Check(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If

                        Case "6", "E"
                            '�����O�X�g���[�N�i�����b�h�j�x�[�X(�o�S)�`�F�b�N
                            If fncLongBaseP4Check(objKtbnStrc, _
                                                    intKtbnStrcSeqNo, _
                                                    strOptionSymbol, _
                                                    strMessageCd) = False Then
                                fncCheckSelectOption = False
                            End If
                            ''RM0906034 2009/08/05 Y.Miura�@�񎟓d�r�Ή��@��ǉ�
                            ''Case "K"
                            'Case "K", "L"
                            '    '���׏d�x�[�X�`�F�b�N
                            '    If fncHighLoadBaseCheck(objKtbnStrc, _
                            '                            intKtbnStrcSeqNo, _
                            '                            strOptionSymbol, _
                            '                            strMessageCd) = False Then
                            '        fncCheckSelectOption = False
                            '    End If
                            'Case "M"
                            '    '���~�߃x�[�X�`�F�b�N
                            '    If fncNonRotatingBaseCheck(objKtbnStrc, _
                            '                            intKtbnStrcSeqNo, _
                            '                            strOptionSymbol, _
                            '                            strMessageCd) = False Then
                            '        fncCheckSelectOption = False
                            '    End If
                            'Case "Q"
                            '    '�����h�~�x�[�X�`�F�b�N
                            '    If fncPositionLockingBaseCheck(objKtbnStrc, _
                            '                            intKtbnStrcSeqNo, _
                            '                            strOptionSymbol, _
                            '                            strMessageCd) = False Then
                            '        fncCheckSelectOption = False
                            '    End If
                            'Case "X"
                            '    '���o���x�[�X�`�F�b�N
                            '    If fncSpringReturnBaseCheck(objKtbnStrc, _
                            '                            intKtbnStrcSeqNo, _
                            '                            strOptionSymbol, _
                            '                            strMessageCd) = False Then
                            '        fncCheckSelectOption = False
                            '    End If
                            'Case "Y"
                            '    '�����݃x�[�X�`�F�b�N
                            '    If fncSpringExtendBaseCheck(objKtbnStrc, _
                            '                            intKtbnStrcSeqNo, _
                            '                            strOptionSymbol, _
                            '                            strMessageCd) = False Then
                            '        fncCheckSelectOption = False
                            '    End If
                            '2010/11/02 MOD RM1011020(12��VerUP:SSD2�V���[�Y) <---END

                    End Select
            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
    '********************************************************************************************
    '*�y�֐����z
    '*  fncStandardBaseCheck
    '*�y�����z
    '*  ��{�x�[�X�`�F�b�N
    '*�y�T�v�z
    '*  ��{�x�[�X���`�F�b�N����
    '*�y�����z
    '*  <Object>       objKtbnStrc          �����`�ԏ��
    '*  <Integer>      intKtbnStrcSeqNo     �`�ԍ\������
    '*  <String>       strOptionSymbol      �I�v�V�����L��
    '*  <String>       strMessageCd         ���b�Z�[�W�R�[�h
    '*�y�߂�l�z
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncStandardBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Try

            fncStandardBaseCheck = True

            '*-----�I�v�V�����`�F�b�N-----*
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1)
                Case "G", "G2", "G3"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4)
                        Case "16", "20", "25", "32"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(20)
                                Case "FA", "LB"
                                    intKtbnStrcSeqNo = 20
                                    strMessageCd = "W9050"
                                    fncStandardBaseCheck = False
                                    Exit Try
                            End Select
                    End Select
            End Select


            '*-----<< �T�D�ŏ��X�g���[�N�`�F�b�N >>-----*
            Dim selList As New ArrayList

            Select Case objKtbnStrc.strcSelection.strOpSymbol(1)
                Case "T1L"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(18)
                        Case "R", "H"
                            selList.Add("10:16,32,40,50,63,80,100")
                            selList.Add("15:20,25")

                        Case "D"
                            selList.Add("20:16,25,32,40,50,63,80,100")
                            selList.Add("25:20")

                    End Select

                    '�X�g���[�N�̃`�F�b�N
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 2) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If

                    '2010/11/02 ADD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                    selList.Clear()
                    selList.Add("50:20,25")

                    '�X�g���[�N�̃`�F�b�N
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                Case ""

                    '�o���G�[�V�����A
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "L4" Then
                        selList.Add("20:*")

                        '�X�g���[�N�̃`�F�b�N
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 2) = False Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If

                    End If

                    selList.Clear()
                    selList.Add("50:20,25")

                    '�X�g���[�N�̃`�F�b�N
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If

                Case "G1"
                    '�o���G�[�V�����A
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "L4" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(18).Trim
                            Case "R", "H", "D"
                                selList.Add("20:*")
                            Case "T"
                                selList.Add("35:*")
                        End Select

                        '�X�g���[�N�̃`�F�b�N
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 2) = False Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If

                    End If

                    selList.Clear()
                    selList.Add("50:20,25")

                    '�X�g���[�N�̃`�F�b�N
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If

                Case "T1", "O", "G", "G2", "G3", "G4", "G5"
                    selList.Add("50:20,25")

                    '�X�g���[�N�̃`�F�b�N
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If
                    '2010/11/02 ADD RM1011020(12��VerUP:SSD2�V���[�Y) <---END

                Case "W"
                    selList.Add("30:12,16")
                    selList.Add("50:20,25,32,40,50,63,80,100")
                    selList.Add("300:125,140,160")

                    '�r�P�X�g���[�N�̗L��
                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> 0 Then
                        '�r�P�X�g���[�N�`�F�b�N
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 1) = False Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If

                    '�r�Q�X�g���[�N�̗L��
                    If objKtbnStrc.strcSelection.strOpSymbol(14) <> 0 Then

                        '�r�Q�̃`�F�b�N
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 1) = False Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    End If
                    '2010/11/02 ADD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                Case "M"

                    selList.Add("5:12,16,20,25,32,40")
                    selList.Add("10:50,63")

                    '�X�g���[�N�̃`�F�b�N
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 2) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncStandardBaseCheck = False
                        Exit Try
                    End If

                Case "X", "Y"
                    '�r�Q�X�g���[�N�`�F�b�N
                    If objKtbnStrc.strcSelection.strOpSymbol(14) <> 0 Then
                        '���a��
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4)
                            Case "12", "16", "20", "25", "32", "40"
                                '�r�Q�`�F�b�N
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "5", "10"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncStandardBaseCheck = False
                                        Exit Try
                                End Select
                            Case "50"
                                '�r�Q�`�F�b�N
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "10", "20"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncStandardBaseCheck = False
                                        Exit Try
                                End Select
                        End Select
                    End If
                Case "Q"
                    '�r�Q�X�g���[�N�`�F�b�N
                    If objKtbnStrc.strcSelection.strOpSymbol(14) <> 0 Then
                        '���a��
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4)
                            Case "20", "25", "32", "40", "50", "63"
                                '�r�Q�`�F�b�N
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "10", "15", "20", "25", "50", "75", "100"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncStandardBaseCheck = False
                                        Exit Try
                                End Select
                            Case "80", "100"
                                '�r�Q�`�F�b�N
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "25", "50", "75", "100"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncStandardBaseCheck = False
                                        Exit Try
                                End Select
                        End Select
                    End If
                    '2010/11/02 ADD RM1011020(12��VerUP:SSD2�V���[�Y) <---END
                Case Else
                    '�`�F�b�N�Ȃ�
            End Select

            '�r�P�X�C�b�`���̃`�F�b�N
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(11).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 7
                strMessageCd = "W0200"
                fncStandardBaseCheck = False
                Exit Try
            End If

            '�r�Q�X�C�b�`���̃`�F�b�N
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(16).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(18).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 14
                strMessageCd = "W0200"
                fncStandardBaseCheck = False
                Exit Try
            End If

            '*-----<< �U�D�ő�X�g���[�N�ƃo���G�[�V�����̃`�F�b�N >>-----*
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "M"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                        Case "20", "25"
                            If CDec(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 50 Then
                                intKtbnStrcSeqNo = 14
                                strMessageCd = "W0200"
                                fncStandardBaseCheck = False
                                Exit Try
                            End If
                        Case Else
                    End Select
                Case Else
            End Select

            '*-----<< �U�D�ő�X�g���[�N�ƃS���N�b�V�����̃`�F�b�N >>-----*
            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(5).Trim, "D") = 0 Then
            Else
                Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                    Case "12", "16"
                        If CDec(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 30 Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    Case "20", "25"
                        If CDec(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 50 Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                    Case "32", "40", "50", "63", "80", "100"
                        If CDec(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) > 100 Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncStandardBaseCheck = False
                            Exit Try
                        End If
                End Select
            End If

            '*-----<< �V�D�ő�X�g���[�N�ƍŏ��X�g���[�N�̑��փ`�F�b�N >>-----*
            '��i�`�̎�
            If objKtbnStrc.strcSelection.strOpSymbol(1) = "W" Then

                If IsNumeric(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) AndAlso _
                IsNumeric(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) AndAlso _
                CInt(objKtbnStrc.strcSelection.strOpSymbol(7).Trim) < CInt(objKtbnStrc.strcSelection.strOpSymbol(14).Trim) Then
                    intKtbnStrcSeqNo = 14
                    strMessageCd = "W0610"
                    fncStandardBaseCheck = False
                    Exit Try

                End If
            End If

            '201012/10 ADD RM1012055(1��VerUP::SSD2�V���[�Y) START--->
            '*-----<< �I�v�V�����u���ԃX�g���[�N��p�{�́v�`�F�b�N >>-----*
            Dim strOp() As String
            strOp = Split(objKtbnStrc.strcSelection.strOpSymbol(19).Trim, ",")
            If Not fncOptionSCheck(strOp, _
                                objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                objKtbnStrc.strcSelection.strOpSymbol(14).Trim) Then
                intKtbnStrcSeqNo = 14
                strMessageCd = "W0830"
                fncStandardBaseCheck = False
                Exit Try
            End If
            '2010/12/10 ADD RM1012055(1��VerUP:SSD2�V���[�Y) <---END

        Catch ex As Exception

            Throw ex

        End Try

    End Function
    '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) <---END

    '********************************************************************************************
    '*�y�֐����z
    '*  fncP4BaseCheck
    '*�y�����z
    '*  �񎟓d�r�x�[�X�`�F�b�N
    '*�y�T�v�z
    '*  �񎟓d�r�x�[�X���`�F�b�N����
    '*�y�����z
    '*  <Object>       objKtbnStrc          �����`�ԏ��
    '*  <Integer>      intKtbnStrcSeqNo     �`�ԍ\������
    '*  <String>       strOptionSymbol      �I�v�V�����L��
    '*  <String>       strMessageCd         ���b�Z�[�W�R�[�h
    '*�y�߂�l�z
    '*  <Boolean>
    '********************************************************************************************
    '2010/10/05 MOD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
    Private Function fncP4BaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean
        'Private Function fncStandardBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
        '2010/10/05 MOD RM1010017(11��VerUP:SSD2�V���[�Y) <---END

        Try

            fncP4BaseCheck = True

            '*-----<< �T�D�ŏ��X�g���[�N�`�F�b�N >>-----*
            '�X�C�b�`���̃`�F�b�N
            'RM1210067 2013/02/01 Y.Tachi ���[�J���łƂ̍��ُC��
            '�r�P
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(9).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(11).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 7
                strMessageCd = "W0200"
                fncP4BaseCheck = False
                Exit Try
            End If
            '�r�Q
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(16).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(18).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 14
                strMessageCd = "W0200"
                fncP4BaseCheck = False
                Exit Try
            End If
            '*-----<< �U�D�ő�X�g���[�N�ƃS���N�b�V�����̃`�F�b�N >>-----*

            If InStr(1, objKtbnStrc.strcSelection.strOpSymbol(4).Trim, "D") = 0 Then
            Else
                Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                    Case "12", "16"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 30 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W0200"
                            fncP4BaseCheck = False
                            Exit Try
                        End If
                    Case "20", "25"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 50 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W0200"
                            fncP4BaseCheck = False
                            Exit Try
                        End If
                    Case "32", "40", "50", "63", "80", "100"
                        If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) > 100 Then
                            intKtbnStrcSeqNo = 5
                            strMessageCd = "W0200"
                            fncP4BaseCheck = False
                            Exit Try
                        End If
                End Select
            End If

            Dim selList As New ArrayList

            Select Case objKtbnStrc.strcSelection.strOpSymbol(1)
                Case "T1L"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(18)
                        Case "R", "H"
                            selList.Add("10:16,32,40,50,63,80,100")
                            selList.Add("15:20,25")

                        Case "D"
                            selList.Add("20:16,25,32,40,50,63,80,100")
                            selList.Add("25:20")

                    End Select

                    '�X�g���[�N�̃`�F�b�N
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 2) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncP4BaseCheck = False
                        Exit Try
                    End If

                    '2010/11/02 ADD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                    selList.Clear()
                    selList.Add("50:20,25")

                    '�X�g���[�N�̃`�F�b�N
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncP4BaseCheck = False
                        Exit Try
                    End If
                Case ""

                    '�o���G�[�V�����A
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "L4" Then
                        selList.Add("20:*")

                        '�X�g���[�N�̃`�F�b�N
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 2) = False Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncP4BaseCheck = False
                            Exit Try
                        End If

                    End If

                    selList.Clear()
                    selList.Add("50:20,25")

                    '�X�g���[�N�̃`�F�b�N
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncP4BaseCheck = False
                        Exit Try
                    End If

                Case "G1"
                    '�o���G�[�V�����A
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "L4" Then
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(18).Trim
                            Case "R", "H", "D"
                                selList.Add("20:*")
                            Case "T"
                                selList.Add("35:*")
                        End Select

                        '�X�g���[�N�̃`�F�b�N
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 2) = False Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncP4BaseCheck = False
                            Exit Try
                        End If

                    End If

                    selList.Clear()
                    selList.Add("50:20,25")

                    '�X�g���[�N�̃`�F�b�N
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncP4BaseCheck = False
                        Exit Try
                    End If

                Case "T1", "O", "G", "G2", "G3", "G4", "G5"
                    selList.Add("50:20,25")

                    '�X�g���[�N�̃`�F�b�N
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 1) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncP4BaseCheck = False
                        Exit Try
                    End If
                    '2010/11/02 ADD RM1011020(12��VerUP:SSD2�V���[�Y) <---END

                Case "W"
                    selList.Add("30:12,16")
                    selList.Add("50:20,25,32,40,50,63,80,100")
                    selList.Add("300:125,140,160")

                    '�r�P�X�g���[�N�̗L��
                    If objKtbnStrc.strcSelection.strOpSymbol(7) <> 0 Then
                        '�r�P�X�g���[�N�`�F�b�N
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(7).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 1) = False Then
                            intKtbnStrcSeqNo = 7
                            strMessageCd = "W0200"
                            fncP4BaseCheck = False
                            Exit Try
                        End If
                    End If

                    '�r�Q�X�g���[�N�̗L��
                    If objKtbnStrc.strcSelection.strOpSymbol(14) <> 0 Then

                        '�r�Q�̃`�F�b�N
                        If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                                  objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                  selList, 1) = False Then
                            intKtbnStrcSeqNo = 14
                            strMessageCd = "W0200"
                            fncP4BaseCheck = False
                            Exit Try
                        End If
                    End If
                    '2010/11/02 ADD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                Case "M"

                    selList.Add("5:12,16,20,25,32,40")
                    selList.Add("10:50,63")

                    '�X�g���[�N�̃`�F�b�N
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                              selList, 2) = False Then
                        intKtbnStrcSeqNo = 14
                        strMessageCd = "W0200"
                        fncP4BaseCheck = False
                        Exit Try
                    End If

                Case "X", "Y"
                    '�r�Q�X�g���[�N�`�F�b�N
                    If objKtbnStrc.strcSelection.strOpSymbol(14) <> 0 Then
                        '���a��
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4)
                            Case "12", "16", "20", "25", "32", "40"
                                '�r�Q�`�F�b�N
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "5", "10"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncP4BaseCheck = False
                                        Exit Try
                                End Select
                            Case "50"
                                '�r�Q�`�F�b�N
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "10", "20"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncP4BaseCheck = False
                                        Exit Try
                                End Select
                        End Select
                    End If
                Case "Q"
                    '�r�Q�X�g���[�N�`�F�b�N
                    If objKtbnStrc.strcSelection.strOpSymbol(14) <> 0 Then
                        '���a��
                        Select Case objKtbnStrc.strcSelection.strOpSymbol(4)
                            Case "20", "25", "32", "40", "50", "63"
                                '�r�Q�`�F�b�N
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "10", "15", "20", "25", "50", "75", "100"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncP4BaseCheck = False
                                        Exit Try
                                End Select
                            Case "80", "100"
                                '�r�Q�`�F�b�N
                                Select Case objKtbnStrc.strcSelection.strOpSymbol(14)
                                    Case "25", "50", "75", "100"
                                    Case Else
                                        intKtbnStrcSeqNo = 14
                                        strMessageCd = "W0200"
                                        fncP4BaseCheck = False
                                        Exit Try
                                End Select
                        End Select
                    End If
                Case Else
                    '�`�F�b�N�Ȃ�
            End Select

            '2010/10/05 DEL RM1010017(11��VerUP:SSD2�V���[�Y) START--->
            ''RM0906034 2009/08/05 Y.Miura�@�񎟓d�r�Ή��@��ǉ�
            'If objKtbnStrc.strcSelection.strKeyKataban.Equals("4") Then
            '2010/10/05 DEL RM1010017(11��VerUP:SSD2�V���[�Y) <---END

            '��{�x�[�X�`�F�b�N�@�񎟓d�r�Ή�
            '��2012/01/05 RM1201XXX intOptionPos(9��19)�ύX
            If fncP4Check(objKtbnStrc, _
                                    intKtbnStrcSeqNo, _
                                    strOptionSymbol, _
                                    strMessageCd, _
                                    19) = False Then
                fncP4BaseCheck = False
                Exit Try
            End If
            'End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*�y�֐����z
    '*  fncDoubleRodBaseCheck
    '*�y�����z
    '*  �����b�h�x�[�X�`�F�b�N
    '*�y�T�v�z
    '*  �����b�h�x�[�X���`�F�b�N����
    '*�y�����z
    '*  <Object>       objKtbnStrc          �����`�ԏ��
    '*  <Integer>      intKtbnStrcSeqNo     �`�ԍ\������
    '*  <String>       strOptionSymbol      �I�v�V�����L��
    '*  <String>       strMessageCd         ���b�Z�[�W�R�[�h
    '*�y�߂�l�z
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncDoubleRodBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Try

            fncDoubleRodBaseCheck = True

            '*-----<< �T�D�ŏ��X�g���[�N�`�F�b�N >>-----*
            '2010/11/02 ADD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
            Dim selList As New ArrayList

            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "D"
                    selList.Add("5:25,32,40")
                    selList.Add("10:50,63,80,100")

                    '�X�g���[�N�`�F�b�N
                    If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
                                              objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                              selList, 2) = False Then
                        intKtbnStrcSeqNo = 5
                        strMessageCd = "W0200"
                        fncDoubleRodBaseCheck = False
                        Exit Try
                    End If
            End Select
            '2010/11/02 ADD RM1011020(12��VerUP:SSD2�V���[�Y) <---END

            '�X�C�b�`���̃`�F�b�N
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 5
                strMessageCd = "W0200"
                fncDoubleRodBaseCheck = False
                Exit Try
            End If

            '*-----<< �U�D���ԃX�g���[�N�`�F�b�N >>-----*

            Select Case objKtbnStrc.strcSelection.strOpSymbol(3).Trim
                Case "32", "40", "50", "63", "80", "100"
                    ' ���ԃX�g���[�N�`�F�b�N(5mm��)
                    If CInt(objKtbnStrc.strcSelection.strOpSymbol(5).Trim) Mod 5 <> 0 Then
                        intKtbnStrcSeqNo = 5
                        strMessageCd = "W0510"
                        fncDoubleRodBaseCheck = False
                    End If

            End Select

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*�y�֐����z
    '*  fncHighLoadBaseCheck
    '*�y�����z
    '*  ���׏d�x�[�X�`�F�b�N
    '*�y�T�v�z
    '*  ���׏d�x�[�X���`�F�b�N����
    '*�y�����z
    '*  <Object>       objKtbnStrc          �����`�ԏ��
    '*  <Integer>      intKtbnStrcSeqNo     �`�ԍ\������
    '*  <String>       strOptionSymbol      �I�v�V�����L��
    '*  <String>       strMessageCd         ���b�Z�[�W�R�[�h
    '*�y�߂�l�z
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncHighLoadBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Try

            fncHighLoadBaseCheck = True

            '*-----�I�v�V�����`�F�b�N-----*
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1)
                Case "KG2", "KG3"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(4)
                        Case "16", "20", "25", "32"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(20)
                                Case "FA", "LB"
                                    intKtbnStrcSeqNo = 20
                                    strMessageCd = "W9050"
                                    fncHighLoadBaseCheck = False
                                    Exit Try
                            End Select
                    End Select
            End Select

            '2010/11/02 ADD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
            '*-----<< �T�D�X�g���[�N�`�F�b�N >>-----*
            Dim selListMin As New ArrayList
            '�o���G�[�V�����@
            Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                Case "K", "KG1"
                    '�z�ǂ˂��A�N�b�V����
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(5).Trim
                        Case "C", "GC", "NC"
                            '���a
                            selListMin.Add("5:20,25,32,40,50")
                            selListMin.Add("10:63,80,100")

                            If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                                      objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                                      selListMin, 2) = False Then
                                intKtbnStrcSeqNo = 14
                                strMessageCd = "W0200"
                                fncHighLoadBaseCheck = False
                                Exit Try
                            End If

                            selListMin.Clear()

                    End Select
            End Select

            '�o���G�[�V�����A
            Select Case objKtbnStrc.strcSelection.strOpSymbol(2).Trim
                Case "L"

                    '�X�C�b�`
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                        Case ""
                            'Do Nothing
                        Case "T0H", "T0V", "T5H", "T5V"
                            '�ŏ��`�F�b�N�l�ݒ�
                            selListMin.Add("10:12,16")
                            selListMin.Add("5:20,25,32,40,50,63,80,100")
                        Case "F2V", "F3V", "F2YV", "F3YV"
                            '�ŏ��`�F�b�N�l�ݒ�
                            selListMin.Add("10:20")
                            selListMin.Add("5:12,16,25,32,40,50,63,80,100")
                        Case "T2H", "T2V", "T3H", "T3V", "T3PH", "T3PV", "F2H", "F2V", "F2YH", "F3YH"
                            '�ŏ��`�F�b�N�l�ݒ�
                            selListMin.Add("5:*")
                        Case Else
                            '�ŏ��`�F�b�N�l�ݒ�
                            selListMin.Add("10:*")
                    End Select

                    '�o���G�[�V�����@
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "K", "KG1"
                            '�X�C�b�`
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                                Case "T0H", "T0V", "T5H", "T5V", "T2H", "T2V", "T3H", "T3V"
                                    '�r�Q�F��
                                    If objKtbnStrc.strcSelection.strOpSymbol(18).Trim = "D" Then
                                        '�ŏ��`�F�b�N�l�ݒ�
                                        selListMin.Clear()  '����L�Őݒ肳��Ă���ꍇ�A�ŏ��`�F�b�N�l��h�ւ���
                                        selListMin.Add("5:*")
                                    End If
                            End Select
                    End Select
                    '��RM1212080 2012/12/05 Y.Tachi
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "K"
                            Select Case objKtbnStrc.strcSelection.strOpSymbol(4).Trim
                                Case "12", "16"
                                    Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                                        Case "T0H", "T0V", "T5H", "T5V"
                                            Select Case objKtbnStrc.strcSelection.strOpSymbol(18).Trim
                                                Case "R"
                                                    '�ŏ��`�F�b�N�l�ݒ�
                                                    selListMin.Clear()  '����L�Őݒ肳��Ă���ꍇ�A�ŏ��`�F�b�N�l��h�ւ���
                                                    selListMin.Add("5:*")
                                                Case "D"
                                                    '�ŏ��`�F�b�N�l�ݒ�
                                                    selListMin.Clear()  '����L�Őݒ肳��Ă���ꍇ�A�ŏ��`�F�b�N�l��h�ւ���
                                                    selListMin.Add("10:*")
                                            End Select
                                    End Select
                            End Select
                    End Select
                    '��RM1212080 2012/12/05 Y.Tachi
                Case "L4"

                    '�X�C�b�`
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(16).Trim
                        Case ""
                            'Do Nothing
                        Case "T0H", "T0V", "T5H", "T5V"
                            '�ŏ��`�F�b�N�l�ݒ�
                            selListMin.Add("10:12,16")
                            selListMin.Add("5:20,25,32,40,50,63,80,100")
                        Case "F2V", "F3V", "F2YV", "F3YV"
                            '�ŏ��`�F�b�N�l�ݒ�
                            selListMin.Add("10:20")
                            selListMin.Add("5:12,16,25,32,40,50,63,80,100")
                        Case "T2H", "T2V", "T3H", "T3V", "T3PH", "T3PV", "F2H", "F2V", "F2YH", "F3YH"
                            '�ŏ��`�F�b�N�l�ݒ�
                            selListMin.Add("5:*")
                        Case Else
                            '�ŏ��`�F�b�N�l�ݒ�
                            selListMin.Add("10:*")
                    End Select

                    '�o���G�[�V�����A
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1).Trim
                        Case "K"
                            '�ŏ��`�F�b�N�l�ݒ�
                            selListMin.Clear()  '����L�Őݒ肳��Ă���ꍇ�A�ŏ��`�F�b�N�l��h�ւ���
                            selListMin.Add("20:*")

                        Case "KG1"
                            '�ŏ��`�F�b�N�l�ݒ�
                            selListMin.Clear()  '����L�Őݒ肳��Ă���ꍇ�A�ŏ��`�F�b�N�l��h�ւ���
                            selListMin.Add("20:*")

                    End Select
            End Select

            '�ŏ��X�g���[�N�`�F�b�N
            If fncSSD2BaseStrokeCheck(objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                          objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                          selListMin, 2) = False Then
                intKtbnStrcSeqNo = 14
                strMessageCd = "W0200"
                fncHighLoadBaseCheck = False
                Exit Try
            End If

            '2010/11/02 ADD RM1011020(12��VerUP:SSD2�V���[�Y) <---END

            '�X�C�b�`���̃`�F�b�N
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 5
                strMessageCd = "W0200"
                fncHighLoadBaseCheck = False
                Exit Try
            End If

            '2010/12/10 ADD RM1012055(1��VerUP:SSD2�V���[�Y) START--->
            '*-----<< �I�v�V�����u���ԃX�g���[�N��p�{�́v�`�F�b�N >>-----*
            Dim strOp() As String
            strOp = Split(objKtbnStrc.strcSelection.strOpSymbol(19).Trim, ",")
            If Not fncOptionSCheck(strOp, _
                                objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                objKtbnStrc.strcSelection.strOpSymbol(14).Trim) Then
                intKtbnStrcSeqNo = 14
                strMessageCd = "W0830"
                fncHighLoadBaseCheck = False
                Exit Try
            End If
            '2010/12/10 ADD RM1012055(1��VerUP:SSD2�V���[�Y) <---END

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*�y�֐����z
    '*  fncHighLoadBaseCheck
    '*�y�����z
    '*  ���׏d�x�[�X�`�F�b�N
    '*�y�T�v�z
    '*  ���׏d�x�[�X���`�F�b�N����
    '*�y�����z
    '*  <Object>       objKtbnStrc          �����`�ԏ��
    '*  <Integer>      intKtbnStrcSeqNo     �`�ԍ\������
    '*  <String>       strOptionSymbol      �I�v�V�����L��
    '*  <String>       strMessageCd         ���b�Z�[�W�R�[�h
    '*�y�߂�l�z
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncHighLoadBaseP4Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Try

            fncHighLoadBaseP4Check = True

            '*-----<< �T�D�ŏ��X�g���[�N�`�F�b�N >>-----*
            '�X�C�b�`���̃`�F�b�N
            '�r�P
            'RM1305005 2013/05/30 ���[�J���łƍ��ُC��
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(14).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(16).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(18).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(4).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 5
                strMessageCd = "W0200"
                fncHighLoadBaseP4Check = False
                Exit Try
            End If

            '2010/11/02 DEL RM1011020(12��VerUP:SSD2�V���[�Y) START--->
            ''RM0906034 2009/08/05 Y.Miura�@�񎟓d�r�Ή��@��ǉ�
            'If objKtbnStrc.strcSelection.strKeyKataban.Equals("L") Then
            '2010/11/02 DEL RM1011020(12��VerUP:SSD2�V���[�Y) <---END

            '��{�x�[�X�`�F�b�N�@�񎟓d�r�Ή�
            '��2012/01/05 RM1201XXX intOptionPos(9��19)�ύX
            If fncP4Check(objKtbnStrc, _
                                    intKtbnStrcSeqNo, _
                                    strOptionSymbol, _
                                    strMessageCd, _
                                    19) = False Then
                fncHighLoadBaseP4Check = False
                Exit Try
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*�y�֐����z
    '*  fncLongBaseP4Check
    '*�y�����z
    '*  �����O�X�g���[�N�i�����b�h�j�i�o�S�j�x�[�X�`�F�b�N
    '*�y�T�v�z
    '*  �����O�X�g���[�N�i�����b�h�j�i�o�S�j���`�F�b�N����
    '*�y�����z
    '*  <Object>       objKtbnStrc          �����`�ԏ��
    '*  <Integer>      intKtbnStrcSeqNo     �`�ԍ\������
    '*  <String>       strOptionSymbol      �I�v�V�����L��
    '*  <String>       strMessageCd         ���b�Z�[�W�R�[�h
    '*�y�߂�l�z
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncLongBaseP4Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String) As Boolean

        Try

            fncLongBaseP4Check = True

            '*-----<< �T�D�ŏ��X�g���[�N�`�F�b�N >>-----*
            '�X�C�b�`���̃`�F�b�N
            If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(3).Trim, _
                                        objKtbnStrc.strcSelection.strOpSymbol(2).Trim) = False Then
                intKtbnStrcSeqNo = 5
                strMessageCd = "W0200"
                fncLongBaseP4Check = False
                Exit Try
            End If

            '��{�x�[�X�`�F�b�N�@�񎟓d�r�Ή�
            If fncP4Check(objKtbnStrc, _
                                    intKtbnStrcSeqNo, _
                                    strOptionSymbol, _
                                    strMessageCd, _
                                    9) = False Then
                fncLongBaseP4Check = False
                Exit Try
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '2010/11/02 DEL RM1011020(12��VerUP:SSD2�V���[�Y) START--->
    ''********************************************************************************************
    ''*�y�֐����z
    ''*  fncNonRotatingBaseCheck
    ''*�y�����z
    ''*  ���~�߃x�[�X�`�F�b�N
    ''*�y�T�v�z
    ''*  ���~�߃x�[�X���`�F�b�N����
    ''*�y�����z
    ''*  <Object>       objKtbnStrc          �����`�ԏ��
    ''*  <Integer>      intKtbnStrcSeqNo     �`�ԍ\������
    ''*  <String>       strOptionSymbol      �I�v�V�����L��
    ''*  <String>       strMessageCd         ���b�Z�[�W�R�[�h
    ''*�y�߂�l�z
    ''*  <Boolean>
    ''********************************************************************************************
    'Private Function fncNonRotatingBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
    '                                      ByRef intKtbnStrcSeqNo As Integer, _
    '                                      ByRef strOptionSymbol As String, _
    '                                      ByRef strMessageCd As String) As Boolean

    '    Try

    '        fncNonRotatingBaseCheck = True

    '        '*-----<< �T�D�ŏ��X�g���[�N�`�F�b�N >>-----*
    '        '�X�C�b�`���̃`�F�b�N
    '        If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = False Then
    '            intKtbnStrcSeqNo = 5
    '            strMessageCd = "W0200"
    '            fncNonRotatingBaseCheck = False
    '            Exit Try
    '        End If

    '    Catch ex As Exception

    '        Throw ex

    '    End Try

    'End Function

    ''********************************************************************************************
    ''*�y�֐����z
    ''*  fncPositionLockingBaseCheck
    ''*�y�����z
    ''*  �����h�~�x�[�X�`�F�b�N
    ''*�y�T�v�z
    ''*  �����h�~�x�[�X���`�F�b�N����
    ''*�y�����z
    ''*  <Object>       objKtbnStrc          �����`�ԏ��
    ''*  <Integer>      intKtbnStrcSeqNo     �`�ԍ\������
    ''*  <String>       strOptionSymbol      �I�v�V�����L��
    ''*  <String>       strMessageCd         ���b�Z�[�W�R�[�h
    ''*�y�߂�l�z
    ''*  <Boolean>
    ''********************************************************************************************
    'Private Function fncPositionLockingBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
    '                                      ByRef intKtbnStrcSeqNo As Integer, _
    '                                      ByRef strOptionSymbol As String, _
    '                                      ByRef strMessageCd As String) As Boolean

    '    Try

    '        fncPositionLockingBaseCheck = True

    '        '*-----<< �T�D�ŏ��X�g���[�N�`�F�b�N >>-----*
    '        '�X�C�b�`���̃`�F�b�N
    '        If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = False Then
    '            intKtbnStrcSeqNo = 5
    '            strMessageCd = "W0200"
    '            fncPositionLockingBaseCheck = False
    '            Exit Try
    '        End If

    '    Catch ex As Exception

    '        Throw ex

    '    End Try

    'End Function

    ''********************************************************************************************
    ''*�y�֐����z
    ''*  fncDoubleRodBaseCheck
    ''*�y�����z
    ''*  ���o���x�[�X�`�F�b�N
    ''*�y�T�v�z
    ''*  ���o���x�[�X���`�F�b�N����
    ''*�y�����z
    ''*  <Object>       objKtbnStrc          �����`�ԏ��
    ''*  <Integer>      intKtbnStrcSeqNo     �`�ԍ\������
    ''*  <String>       strOptionSymbol      �I�v�V�����L��
    ''*  <String>       strMessageCd         ���b�Z�[�W�R�[�h
    ''*�y�߂�l�z
    ''*  <Boolean>
    ''********************************************************************************************
    'Private Function fncSpringReturnBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
    '                                      ByRef intKtbnStrcSeqNo As Integer, _
    '                                      ByRef strOptionSymbol As String, _
    '                                      ByRef strMessageCd As String) As Boolean

    '    Try

    '        fncSpringReturnBaseCheck = True

    '        '*-----<< �T�D�ŏ��X�g���[�N�`�F�b�N >>-----*
    '        '�X�C�b�`���̃`�F�b�N
    '        If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = False Then
    '            intKtbnStrcSeqNo = 5
    '            strMessageCd = "W0200"
    '            fncSpringReturnBaseCheck = False
    '            Exit Try
    '        End If

    '    Catch ex As Exception

    '        Throw ex

    '    End Try

    'End Function

    ''********************************************************************************************
    ''*�y�֐����z
    ''*  fncSpringExtendBaseCheck
    ''*�y�����z
    ''*  �����݃x�[�X�`�F�b�N
    ''*�y�T�v�z
    ''*  �����݃x�[�X���`�F�b�N����
    ''*�y�����z
    ''*  <Object>       objKtbnStrc          �����`�ԏ��
    ''*  <Integer>      intKtbnStrcSeqNo     �`�ԍ\������
    ''*  <String>       strOptionSymbol      �I�v�V�����L��
    ''*  <String>       strMessageCd         ���b�Z�[�W�R�[�h
    ''*�y�߂�l�z
    ''*  <Boolean>
    ''********************************************************************************************
    'Private Function fncSpringExtendBaseCheck(ByVal objKtbnStrc As KHKtbnStrc, _
    '                                      ByRef intKtbnStrcSeqNo As Integer, _
    '                                      ByRef strOptionSymbol As String, _
    '                                      ByRef strMessageCd As String) As Boolean

    '    Try

    '        fncSpringExtendBaseCheck = True

    '        '*-----<< �T�D�ŏ��X�g���[�N�`�F�b�N >>-----*
    '        '�X�C�b�`���̃`�F�b�N
    '        If fncSSD2SwitchStrokeCheck(objKtbnStrc.strcSelection.strKeyKataban.Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(5).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(6).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(8).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(1).Trim, _
    '                                    objKtbnStrc.strcSelection.strOpSymbol(3).Trim) = False Then
    '            intKtbnStrcSeqNo = 5
    '            strMessageCd = "W0200"
    '            fncSpringExtendBaseCheck = False
    '            Exit Try
    '        End If

    '    Catch ex As Exception

    '        Throw ex

    '    End Try

    'End Function
    '2010/11/02 DEL RM1011020(12��VerUP:SSD2�V���[�Y) <---END

    '********************************************************************************************
    '*�y�֐����z
    '*  fncP4Check
    '*�y�����z
    '*  �񎟓d�r�Ή��@��`�F�b�N
    '*�y�T�v�z
    '*  �񎟓d�r���܂܂�邩���`�F�b�N����
    '*�y�����z
    '*  <Object>       objKtbnStrc          �����`�ԏ��
    '*  <Integer>      intKtbnStrcSeqNo     �`�ԍ\������
    '*  <String>       strOptionSymbol      �I�v�V�����L��
    '*  <String>       strMessageCd         ���b�Z�[�W�R�[�h
    '*  <Integer>      intOptionPos         �v�f�ʒu�@�@�@�@�@   
    '*�y�߂�l�z
    '*  <Boolean>
    '*�y�X�V�z
    '*  �E��tNo�FRM0906034  �񎟓d�r�Ή��@��Ή��@�V�K�ǉ�
    '*                                      �X�V���F2009/09/08   �X�V�ҁFY.Miura
    '********************************************************************************************
    Private Function fncP4Check(ByVal objKtbnStrc As KHKtbnStrc, _
                                          ByRef intKtbnStrcSeqNo As Integer, _
                                          ByRef strOptionSymbol As String, _
                                          ByRef strMessageCd As String, _
                                          ByVal intOptionPos As Integer) As Boolean

        Try

            fncP4Check = True

            '�񎟓d�r�Ή�
            Dim bolOpP4 As Boolean = False
            Dim strOpArray() As String
            strOpArray = Split(objKtbnStrc.strcSelection.strOpSymbol(intOptionPos).Trim, CdCst.Sign.Delimiter.Comma)
            For intLoopCnt As Integer = 0 To strOpArray.Length - 1
                Select Case strOpArray(intLoopCnt).Trim
                    Case "P4", "P40"
                        bolOpP4 = True
                End Select
            Next
            'P4�̕K�{�`�F�b�N
            If Not bolOpP4 Then
                intKtbnStrcSeqNo = intOptionPos
                strMessageCd = "W8770"
                fncP4Check = False
                Exit Try
            End If

        Catch ex As Exception

            Throw ex

        End Try

    End Function

    '********************************************************************************************
    '*�y�֐����z
    '*  fncSSD2SwitchStrokeCheck
    '*�y�����z
    '*  �X�C�b�`���̃`�F�b�N
    '*�y�T�v�z
    '*  �X�C�b�`�`�Ԗ��ɃX�g���[�N���`�F�b�N����
    '*�y�����z
    '*  <String>        strStroke           �X�g���[�N
    '*  <String>        strSwitchKataban    �X�C�b�`�`��
    '*  <String>        strSwitchQty        �X�C�b�`��
    '*  <String>        strVariation        �o���G�[�V����
    '*  <String>        strPortSize         ���a
    '*�y�߂�l�z
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncSSD2SwitchStrokeCheck(ByVal strKeyKataban As String, _
                                              ByVal strStroke As String, _
                                              ByVal strSwitchKataban As String, _
                                              ByVal strSwitchQty As String, _
                                              ByVal strVariation As String, _
                                              ByVal strPortSize As String, _
                                              ByVal strSwitch As String)

        Dim objPrice As New KHUnitPrice

        Try

            fncSSD2SwitchStrokeCheck = False

            '��RM1212080 2012/12/04 Y.Tachi 
            If strKeyKataban = "4" Then
                If InStr(1, strSwitch, "L") = 0 Then
                Else
                    If strSwitchKataban.Length = 0 Then
                    Else
                        Select Case strSwitchKataban.Trim
                            Case "SW17", "SW20", "SW27", "SW28", "SW29", "SW30", "SW69", "SW70", "SWAK"
                                Select Case strPortSize.Trim
                                    Case "12", "16"
                                        Select Case strVariation
                                            Case "", "X", "Y", "O", "B", "W", "M"
                                                If strSwitchQty.Trim = "R" Then
                                                    If Val(strStroke) < 5 Then
                                                        Exit Try
                                                    End If
                                                Else
                                                    If Val(strStroke) < 10 Then
                                                        Exit Try
                                                    End If
                                                End If
                                        End Select
                                End Select
                        End Select
                    End If
                End If
            End If
            If strKeyKataban = "L" Then
                If InStr(1, strSwitch, "L") = 0 Then
                Else
                    If strSwitchKataban.Length = 0 Then
                    Else
                        Select Case strSwitchKataban.Trim
                            Case "SW17", "SW20", "SW27", "SW28", "SW29", "SW30", "SW69", "SW70", "SWAK"
                                Select Case strPortSize.Trim
                                    Case "12", "16"
                                        If strVariation.Trim = "K" And _
                                           strSwitchQty.Trim = "R" Then
                                            If Val(strStroke) < 5 Then
                                                Exit Try
                                            End If
                                        Else
                                            If Val(strStroke) < 10 Then
                                                Exit Try
                                            End If
                                        End If
                                End Select
                        End Select
                    End If
                End If
            End If
            '��RM1212080 2012/12/04 Y.Tachi

            'SW�I��L������
            If strSwitchKataban.Trim = "" Then
            Else
                Select Case strSwitchKataban.Trim
                    'RM0906034 2009/08/05 Y.Miura�@�񎟓d�r�Ή�
                    'Case "T0H", "T0V", "T5H", "T5V"
                    Case "T0H", "T0V", "T5H", "T5V", "SW27"
                        Select Case strPortSize.Trim
                            Case "12", "16"
                                Select Case strKeyKataban
                                    'RM0906034 2009/08/05 Y.Miura�@�񎟓d�r�Ή�
                                    'Case "", "K", "X", "Y"
                                    Case "", "K", "X", "Y", "4"
                                        Select Case strVariation.Trim
                                            Case "", "X", "Y", "O", "B", "W", "M"
                                                If strSwitchQty.Trim = "R" Then
                                                    If Val(strStroke) < 5 Then
                                                        Exit Try
                                                    End If
                                                Else
                                                    If Val(strStroke) < 10 Then
                                                        Exit Try
                                                    End If
                                                End If
                                            Case Else
                                                If CInt(strStroke) < 10 Then
                                                    Exit Try
                                                End If
                                        End Select
                                        '��RM1212080 2012/12/04 Y.Tachi
                                    Case "K"
                                        If strVariation.Trim = "K" And _
                                           strSwitchQty.Trim = "R" Then
                                            If Val(strStroke) < 5 Then
                                                Exit Try
                                            End If
                                        Else
                                            If Val(strStroke) < 5 Then
                                                Exit Try
                                            End If
                                        End If
                                    Case "D"
                                        If strVariation.Trim = "DM" Then
                                            If strSwitchQty.Trim = "R" Then
                                                If Val(strStroke) < 5 Then
                                                    Exit Try
                                                End If
                                            Else
                                                If Val(strStroke) < 10 Then
                                                    Exit Try
                                                End If
                                            End If
                                        End If
                                        '��RM1212080 2012/12/04 Y.Tachi
                                End Select
                            Case Else
                                If CInt(strStroke) < 5 Then
                                    Exit Try
                                End If
                        End Select
                        'RM0906034 2009/08/05 Y.Miura�@�񎟓d�r�Ή�
                        'Case "F2V", "F3V", "F2YV", "F3YV"
                    Case "F2V", "F3V", "F2YV", "F3YV", "SW83", "SW84", "SW87", "SW88"
                        Select Case strPortSize.Trim
                            Case "20"
                                Select Case strKeyKataban
                                    Case "", "D", "M", "X", "Y", "E"
                                        If CInt(strStroke) < 15 Then
                                            Exit Try
                                        End If
                                        'RM0906034 2009/08/05 Y.Miura�@�񎟓d�r�Ή�
                                        'Case "K"
                                    Case "K", "L"
                                        If CInt(strStroke) < 10 Then
                                            Exit Try
                                        End If
                                End Select
                            Case Else
                                If CInt(strStroke) < 5 Then
                                    Exit Try
                                End If
                        End Select
                        'RM0906034 2009/08/05 Y.Miura�@�񎟓d�r�Ή�
                        'Case "T2H", "T2V", "T3H", "T3V", "T3PH", "T3PV", "F2H", "F3H", "F2YH", "F3YH"
                        'RM1210067 2013/02/01 Y.Tachi ���[�J���łƂ̍��ُC��
                        'RM1305005 2013/05/30 ���[�J���łƂ̍��ُC��
                    Case "T2H", "T2V", "T3H", "T3V", "T3PH", "T3PV", "F2H", "F3H", "F2YH", "F3YH", "F3PH", "F3PV", _
                        "SW11", "SW12", "SW13", "SW14", "SW15", "SW16", "SW21", "SW22", "SW23", "SW24", "SW25", "SW26", "SW27", _
                        "SW81", "SW82", "SW83", "SW84", "SW85", "SW86", "SW87", "SW88"
                        If CInt(strStroke) < 5 Then
                            Exit Try
                        End If
                    Case Else
                        If CInt(strStroke) < 10 Then
                            Exit Try
                        End If
                End Select
            End If

            fncSSD2SwitchStrokeCheck = True

        Catch ex As Exception

            Throw ex

        Finally

            objPrice = Nothing

        End Try

    End Function

    '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) START--->
    '********************************************************************************************
    '*�y�֐����z
    '*  fncSSD2BaseStrokeCheck
    '*�y�����z
    '*  �X�C�b�`���̃`�F�b�N(��{�x�[�X�p)
    '*�y�T�v�z
    '*  �X�C�b�`�`�Ԗ��ɃX�g���[�N���`�F�b�N����
    '*�y�����z
    '*  <String>        strStroke           �X�g���[�N
    '*  <String>        strPortSize         ���a
    '*  <String>        lstCheck�@�@�@�@�@�@�`�F�b�N���X�g
    '*  <Integer>       checkFlg�@�@�@�@�@�@�`�F�b�N�t���O(1:���傫���A2:��菬����)
    '*�y�߂�l�z
    '*  <Boolean>
    '********************************************************************************************
    Private Function fncSSD2BaseStrokeCheck(ByVal strStroke As String, _
                                            ByVal strPortSize As String, _
                                            ByVal lstCheck As ArrayList, _
                                            ByVal checkFlg As Integer)

        Dim wkSp() As String

        Try

            fncSSD2BaseStrokeCheck = False

            If strStroke.Equals(String.Empty) Then
                Exit Try
            End If

            For i As Integer = 0 To lstCheck.Count - 1
                '����
                wkSp = Split(lstCheck.Item(i).ToString, ":")

                '���a
                '2010/11/02 MOD RM1011020(12��VerUP:SSD2�V���[�Y) START--->
                If InStr(wkSp(1), strPortSize) > 0 Or wkSp(1) = "*" Then
                    'If InStr(wkSp(1), strPortSize) > 0 Then
                    '2010/11/02 MOD RM1011020(12��VerUP:SSD2�V���[�Y) <---END
                    '�X�g���[�N
                    Select Case checkFlg
                        Case 1
                            If strStroke > CInt(wkSp(0)) Then
                                Exit Try
                            End If
                        Case 2
                            If strStroke < CInt(wkSp(0)) Then
                                Exit Try
                            End If

                    End Select
                End If
            Next

            fncSSD2BaseStrokeCheck = True

        Catch ex As Exception
            Throw ex

        End Try
    End Function
    '2010/10/05 ADD RM1010017(11��VerUP:SSD2�V���[�Y) <---END
    '2010/12/10 ADD RM1012055(1��VerUP:SSD2�V���[�Y) START--->
    ''' <summary>
    ''' �I�v�V�����u���ԃX�g���[�N��p�{�́v�`�F�b�N
    ''' </summary>
    ''' <param name="strOp">�I�v�V�������X�g</param>
    ''' <param name="strBoreSize">���a</param>
    ''' <param name="strStroke">S2�X�g���[�N</param>
    ''' <returns>True:�����AFalse:���s</returns>
    ''' <remarks></remarks>
    Private Function fncOptionSCheck(ByVal strOp() As String, _
                                    ByVal strBoreSize As String, ByVal strStroke As String) As Boolean
        Try
            Dim ret As Boolean = True
            For i As Integer = 0 To strOp.Length - 1
                Select Case Trim(strOp(i))
                    Case "S"
                        Select Case strBoreSize
                            Case "12", "16"
                                Select Case strStroke
                                    Case "5", "10", "15", "20", "25", "30"
                                        ret = False
                                        Exit For
                                End Select
                            Case "20", "25"
                                Select Case strStroke
                                    Case "5", "10", "15", "20", "25", "30", "35", "40", "45", "50"
                                        ret = False
                                        Exit For
                                End Select
                            Case "32", "40"
                                Select Case strStroke
                                    Case "5", "10", "15", "20", "25", "30", "35", "40", "45", "50", "75", "100"
                                        ret = False
                                        Exit For
                                End Select
                            Case "50", "63", "80", "100"
                                Select Case strStroke
                                    Case "10", "15", "20", "25", "30", "35", "40", "45", "50", "75", "100"
                                        ret = False
                                        Exit For
                                End Select
                        End Select
                End Select
            Next

            Return ret

        Catch ex As Exception
            Throw ex

        End Try

    End Function
    '2010/12/10 ADD RM1012055(1��VerUP:SSD2�V���[�Y) <---END
End Module
