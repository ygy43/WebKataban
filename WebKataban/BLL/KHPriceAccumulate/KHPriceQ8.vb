'************************************************************************************
'*  ProgramID  �FKHPriceQ8
'*  Program��  �F�P���v�Z�T�u���W���[��
'*
'*                                      �쐬���F2009/08/11   �쐬�ҁFY.Miura
'*                                      �X�V���F             �X�V�ҁF
'*
'*  �T�v       �FESSD,ELCR�V���[�Y  (�d���A�N�`���G�[�^)
'*
'************************************************************************************
Module KHPriceQ8

    Public Sub subPriceCalculation(ByVal objKtbnStrc As KHKtbnStrc, _
                                   ByRef strOpRefKataban() As String, _
                                   ByRef decOpAmount() As Decimal)

        Try

            '�z���`
            ReDim strOpRefKataban(0)
            ReDim decOpAmount(0)

            '��{���i�L�[
            'RM1312084 2013/12/25
            'RM1402099 2014/02/25 ETS�V���[�Y�ǉ�
            Select Case objKtbnStrc.strcSelection.strSeriesKataban.Trim

                Case "ETV"
                    'RM1410045
                    If objKtbnStrc.strcSelection.strKeyKataban = "T" Or objKtbnStrc.strcSelection.strKeyKataban = "U" Or _
                        objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                        'TOYO�i
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   "T" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���_�Z���T
                        If objKtbnStrc.strcSelection.strOpSymbol(8) <> "N" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                        "T" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(10) <> "N" And _
                            objKtbnStrc.strcSelection.strOpSymbol(10) <> "D" Then
                            '�O���[�X�j�b�v��
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                        "T" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(11) <> "N" Then
                            '��Q�I�v�V�����ǉ�  2017/03/22 �ǉ�
                            '�{�f�B�T�C�Y���u05�v�u06�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ�������
                            If objKtbnStrc.strcSelection.strOpSymbol(1) <> "05" And _
                               objKtbnStrc.strcSelection.strOpSymbol(1) <> "06" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-10"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-05"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                            '�H�i
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  2017/03/22 �C��
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(12)
                            'strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                            '                                           objKtbnStrc.strcSelection.strOpSymbol(11)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    Else
                        '���{�i
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���_�Z���T
                        If objKtbnStrc.strcSelection.strOpSymbol(8) <> "N" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(10) <> "N" And _
                            objKtbnStrc.strcSelection.strOpSymbol(10) <> "D" Then
                            '�O���[�X�j�b�v��
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(11) <> "N" Then
                            '��Q�I�v�V�����ǉ�  RM1702018  2017/02/13 �ǉ�
                            '�{�f�B�T�C�Y���u05�v�u06�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ�������
                            If objKtbnStrc.strcSelection.strOpSymbol(1) <> "05" And _
                               objKtbnStrc.strcSelection.strOpSymbol(1) <> "06" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-10"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-05"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                            '�H�i
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(12)
                            'strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                            '                                           objKtbnStrc.strcSelection.strOpSymbol(11)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                Case "ECS"
                    If objKtbnStrc.strcSelection.strKeyKataban = "T" Or objKtbnStrc.strcSelection.strKeyKataban = "U" Or _
                        objKtbnStrc.strcSelection.strKeyKataban = "V" Or objKtbnStrc.strcSelection.strKeyKataban = "W" Or _
                        objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                        'TOYO�i
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   "T" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���[�^��t���@
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   "T" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1)
                        decOpAmount(UBound(decOpAmount)) = 1

                        If objKtbnStrc.strcSelection.strOpSymbol(8) <> "N" Then
                            '���_�Z���T
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       "T" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(10) <> "N" Then
                            '�O���[�X�j�b�v��
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       "T" & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(11) <> "N" Then
                            '��Q�I�v�V�����ǉ�  2017/03/22 �ǉ�
                            '�{�f�B�T�C�Y���u05�v�u06�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ�������
                            If objKtbnStrc.strcSelection.strOpSymbol(1) <> "05" And _
                               objKtbnStrc.strcSelection.strOpSymbol(1) <> "06" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-10"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-05"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "V" Or objKtbnStrc.strcSelection.strKeyKataban = "W" Then
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  2017/03/22 �C��
                            If objKtbnStrc.strcSelection.strOpSymbol(12) <> "N" Then
                                '�h�K����
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(12)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  2017/03/22 �C��
                            If objKtbnStrc.strcSelection.strOpSymbol(13) <> String.Empty Then
                                '�񎟓d�r
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(13)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                            '�H�i
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  2017/03/22 �C��
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(12)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    Else
                        '���{�i
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3)
                        decOpAmount(UBound(decOpAmount)) = 1

                        If objKtbnStrc.strcSelection.strOpSymbol(4) <> "E" And _
                            objKtbnStrc.strcSelection.strOpSymbol(4) <> "B" Then
                            '���[�^��t���@
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(4) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(8) <> "N" Then
                            '���_�Z���T
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(8)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(10) <> "N" Then
                            '�O���[�X�j�b�v��
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(10)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strOpSymbol(11) <> "N" Then
                            '��Q�I�v�V�����ǉ�  RM1702018  2017/02/13 �ǉ�
                            '�{�f�B�T�C�Y���u05�v�u06�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ�������
                            If objKtbnStrc.strcSelection.strOpSymbol(1) <> "05" And _
                               objKtbnStrc.strcSelection.strOpSymbol(1) <> "06" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-10"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(11) & "-05"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "4" Then
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                            If objKtbnStrc.strcSelection.strOpSymbol(12) <> "N" Then
                                '�h�K����
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(12)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                            If objKtbnStrc.strcSelection.strOpSymbol(13) <> String.Empty Then
                                '�񎟓d�r
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(13)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                            '�H�i
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(12)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If
                    End If
                Case "ETS"
                    'RM1402053
                    If objKtbnStrc.strcSelection.strKeyKataban = "T" Or objKtbnStrc.strcSelection.strKeyKataban = "U" Or _
                        objKtbnStrc.strcSelection.strKeyKataban = "V" Or objKtbnStrc.strcSelection.strKeyKataban = "W" Or _
                         objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���[�^��t���@�ƃX�g���[�N�Ō��Z(�{�f�B�T�C�Y�F13,14,17)
                        If objKtbnStrc.strcSelection.strOpSymbol(1) = "13" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(1) = "14" Then

                            If objKtbnStrc.strcSelection.strOpSymbol(4) = "D" Then

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(4) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3)
                                decOpAmount(UBound(decOpAmount)) = 1

                            End If

                        End If
                        If objKtbnStrc.strcSelection.strOpSymbol(1) = "17" Then

                            If objKtbnStrc.strcSelection.strOpSymbol(4) = "D" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(4) = "R" Or _
                            objKtbnStrc.strcSelection.strOpSymbol(4) = "L" Then

                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(4) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3)
                                decOpAmount(UBound(decOpAmount)) = 1

                            End If

                        End If

                        '���_�Z���T
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '�O���[�X�j�b�v��
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(10) & "-OP"
                        decOpAmount(UBound(decOpAmount)) = 1

                        '��Q�I�v�V�����ǉ�  2017/03/22 �ǉ�
                        '�{�f�B�T�C�Y���u05�v�u06�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ�������
                        If objKtbnStrc.strcSelection.strOpSymbol(1) <> "05" And _
                           objKtbnStrc.strcSelection.strOpSymbol(1) <> "06" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(11) & "-10"
                            decOpAmount(UBound(decOpAmount)) = 1

                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(11) & "-05"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "V" Or objKtbnStrc.strcSelection.strKeyKataban = "W" Then
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  2017/03/22 �C��
                            If objKtbnStrc.strcSelection.strOpSymbol(12) <> "N" Then
                                '�h�K����
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(12)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            If objKtbnStrc.strcSelection.strOpSymbol(13) <> String.Empty Then
                                '�񎟓d�r
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(13)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                            '�H�i
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  2017/03/22 �C��
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(12)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    ElseIf objKtbnStrc.strcSelection.strKeyKataban = "A" Or objKtbnStrc.strcSelection.strKeyKataban = "B" Or _
                           objKtbnStrc.strcSelection.strKeyKataban = "C" Or objKtbnStrc.strcSelection.strKeyKataban = "D" Then
                        '���{Multi Axis�V���[�Y
                        '��{���i
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & Left(objKtbnStrc.strcSelection.strOpSymbol(2), 1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3) & objKtbnStrc.strcSelection.strOpSymbol(4)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���_�E���~�b�g�Z���T
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7)
                        decOpAmount(UBound(decOpAmount)) = 1

                    ElseIf objKtbnStrc.strcSelection.strKeyKataban = "I" Or objKtbnStrc.strcSelection.strKeyKataban = "J" Or _
                     objKtbnStrc.strcSelection.strKeyKataban = "K" Or objKtbnStrc.strcSelection.strKeyKataban = "L" Or _
                     objKtbnStrc.strcSelection.strKeyKataban = "M" Or objKtbnStrc.strcSelection.strKeyKataban = "N" Or _
                     objKtbnStrc.strcSelection.strKeyKataban = "O" Or objKtbnStrc.strcSelection.strKeyKataban = "P" Then
                        '���{Multi Axis�V���[�Y
                        '��{���i
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & Left(objKtbnStrc.strcSelection.strOpSymbol(2), 1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3) & objKtbnStrc.strcSelection.strOpSymbol(4)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���_�E���~�b�g�Z���T
                        If objKtbnStrc.strcSelection.strOpSymbol(1).Trim = "210" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7) & CdCst.Sign.Hypen & "210"
                            decOpAmount(UBound(decOpAmount)) = 1
                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(7)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                    Else
                        '���{�W���i
                        '��{���i
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(3)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���[�^��t���@
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(4) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '���_�Z���T
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8)
                        decOpAmount(UBound(decOpAmount)) = 1

                        '�O���[�X�j�b�v��
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(10) & "-OP"
                        decOpAmount(UBound(decOpAmount)) = 1

                        '��Q�I�v�V�����ǉ�  RM1702018  2017/02/13 �ǉ�
                        '�{�f�B�T�C�Y���u05�v�u06�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ�������
                        If objKtbnStrc.strcSelection.strOpSymbol(1) <> "05" And _
                           objKtbnStrc.strcSelection.strOpSymbol(1) <> "06" Then
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(11) & "-10"
                            decOpAmount(UBound(decOpAmount)) = 1

                        Else
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(11) & "-05"
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "4" Then
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                            If objKtbnStrc.strcSelection.strOpSymbol(12) <> "N" Then
                                '�h�K����
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(12)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If

                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                            If objKtbnStrc.strcSelection.strOpSymbol(13) <> String.Empty Then
                                '�񎟓d�r
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(13)
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If

                        If objKtbnStrc.strcSelection.strKeyKataban = "F" Then
                            '�H�i
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(12)
                            decOpAmount(UBound(decOpAmount)) = 1
                        End If

                        'RM1802016  ���^���X�g�p�d�l�ǉ�
                        '�{�f�B�T�C�Y���u12�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ�������
                        If objKtbnStrc.strcSelection.strKeyKataban = "Z" Then
                            If objKtbnStrc.strcSelection.strOpSymbol(1) = "12" Then
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(12) & "-12"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Else
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(12) & "-06"
                                decOpAmount(UBound(decOpAmount)) = 1
                            End If
                        End If
                    End If

                Case "ECV"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3)
                    decOpAmount(UBound(decOpAmount)) = 1


                    '���_�Z���T
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(8)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�O���[�X�j�b�v��
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(10)
                    decOpAmount(UBound(decOpAmount)) = 1


                    '�ΏۊO�ƂȂ��Ă���L�[�^�Ԃɂ��Ă��ΏۂƂȂ邽�ߏC��  2017/03/22 �C��
                    'If objKtbnStrc.strcSelection.strKeyKataban = "T" Or objKtbnStrc.strcSelection.strKeyKataban = "U" _
                    '   Or objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                    
                    '��Q�I�v�V�����ǉ�  RM1702018  2017/02/13 �ǉ�
                    '�{�f�B�T�C�Y���u05�v�u06�v�̂Ƃ��Ƃ���ȊO�̂Ƃ��ŏ�������
                    If objKtbnStrc.strcSelection.strOpSymbol(1) <> "05" And _
                           objKtbnStrc.strcSelection.strOpSymbol(1) <> "06" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(11) & "-10"
                        decOpAmount(UBound(decOpAmount)) = 1

                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(11) & "-05"
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    
                    If objKtbnStrc.strcSelection.strKeyKataban = "F" Or _
                         objKtbnStrc.strcSelection.strKeyKataban = "X" Or objKtbnStrc.strcSelection.strKeyKataban = "Y" Then
                        '�H�i
                        '�I�v�V�����ǉ��ɂ��A�����邽�ߏC��  RM1702018  2017/02/13 �C��
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(12)
                        'strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                        '                                           objKtbnStrc.strcSelection.strOpSymbol(11)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "ESM"
                    Select Case objKtbnStrc.strcSelection.strOpSymbol(1)
                        Case "HDU", "TTU", "CA", "SE", "PP1"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1)
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "VC"
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(5)
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "ST"
                            Dim intST As Integer = 0
                            intST = Math.Ceiling(objKtbnStrc.strcSelection.strOpSymbol(2) * 0.01) * 100
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                       intST
                            decOpAmount(UBound(decOpAmount)) = 1
                        Case "B"
                            Dim intST As Integer = 0
                            intST = Math.Ceiling(objKtbnStrc.strcSelection.strOpSymbol(3) * 0.01) * 100
                            ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                            ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                            strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                       objKtbnStrc.strcSelection.strOpSymbol(1)
                            decOpAmount(UBound(decOpAmount)) = intST
                    End Select

                Case "ERL"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & "S-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�I�v�V�������Z���i�L�[
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "A" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(9)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "ERL2"
                    '��{���Z���i�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & "E-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '���[�^��t�������Z�L�[
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "E" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                             objKtbnStrc.strcSelection.strOpSymbol(1) & objKtbnStrc.strcSelection.strOpSymbol(2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '�u���[�L���Z���i�L�[
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '�I�v�V�������Z���i�L�[
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "N0" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "N" Then
                    Else

                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                            Case "A", "B"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9) & "-EC07"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Case "C", "D"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9) & "-EC63"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Case "E", "F"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9) & "-ECPT"
                                decOpAmount(UBound(decOpAmount)) = 1
                        End Select

                    End If

                Case "ESD"
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & "S-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4) & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(5)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�I�v�V�������Z���i�L�[
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "A" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(9)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(10)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case "ESD2"
                    '��{���Z���i�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & "E-" & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(4)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '���[�^��t�������Z�L�[
                    If objKtbnStrc.strcSelection.strOpSymbol(2).Trim = "E" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                             objKtbnStrc.strcSelection.strOpSymbol(1) & objKtbnStrc.strcSelection.strOpSymbol(2)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '�u���[�L���Z���i�L�[
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    '�I�v�V�������Z���i�L�[
                    If objKtbnStrc.strcSelection.strOpSymbol(7).Trim = "N0" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(7)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(8).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(8)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
                    If objKtbnStrc.strcSelection.strOpSymbol(9).Trim = "N" Then
                    Else

                        Select Case objKtbnStrc.strcSelection.strOpSymbol(8).Trim
                            Case "A", "B"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9) & "-EC07"
                                decOpAmount(UBound(decOpAmount)) = 1

                            Case "C", "D"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9) & "-EC63"
                                decOpAmount(UBound(decOpAmount)) = 1
                            Case "E", "F"
                                ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                                ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                                strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                           CdCst.Sign.Hypen & _
                                                                           objKtbnStrc.strcSelection.strOpSymbol(9) & "-ECPT"
                                decOpAmount(UBound(decOpAmount)) = 1

                        End Select

                    End If

                    If objKtbnStrc.strcSelection.strOpSymbol(10).Trim = "N" Then
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(10) & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(1)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                    'RM1803042_EBS�EEBR�ǉ�
                Case "EBS"

                    '��{���i���Z�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(4) & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(8)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�Z���T��t
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(9)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�Z���T��t
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(10)
                    decOpAmount(UBound(decOpAmount)) = 1

                Case "EBR"

                    '��{���i���Z�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(5) & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(9)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�Z���T��t
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(10)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�Z���T��t
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(11)
                    decOpAmount(UBound(decOpAmount)) = 1

                    'RM1804032_EKS�ǉ�
                Case "EKS"

                    '��{���i���Z�L�[
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(1) & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(2) & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(3) & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(5)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '���[�^��t���@
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(8) & _
                                                                objKtbnStrc.strcSelection.strOpSymbol(10)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�Z���T��t
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim = "005" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(5) & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(11)
                        decOpAmount(UBound(decOpAmount)) = 1
                    Else
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & CdCst.Sign.Hypen & _
                                                                    objKtbnStrc.strcSelection.strOpSymbol(11)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If

                Case Else
                    ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                    ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                    strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                               CdCst.Sign.Hypen & "BASE" & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(1) & CdCst.Sign.Hypen & _
                                                               objKtbnStrc.strcSelection.strOpSymbol(3)
                    decOpAmount(UBound(decOpAmount)) = 1

                    '�I�v�V�������Z���i�L�[
                    If objKtbnStrc.strcSelection.strOpSymbol(5).Trim <> "" Then
                        ReDim Preserve strOpRefKataban(UBound(strOpRefKataban) + 1)
                        ReDim Preserve decOpAmount(UBound(decOpAmount) + 1)
                        strOpRefKataban(UBound(strOpRefKataban)) = objKtbnStrc.strcSelection.strSeriesKataban.Trim & _
                                                                   CdCst.Sign.Hypen & "OP" & CdCst.Sign.Hypen & _
                                                                   objKtbnStrc.strcSelection.strOpSymbol(5)
                        decOpAmount(UBound(decOpAmount)) = 1
                    End If
            End Select

        Catch ex As Exception

            Throw ex

        Finally

        End Try

    End Sub

End Module

