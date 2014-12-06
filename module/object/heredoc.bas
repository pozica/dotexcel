Attribute VB_Name = "object_heredoc"
'---------------------------------------------------------------------------------------
' Module    : modHereDoc
' Version   : 1.0.0
' DateTime  : 2010/03/27 10:41
' Author    : YU-TANG
' Purpose   : VBA �p�q�A�h�L�������g
' Reference : http://www.f3.dion.ne.jp/~element/msaccess/AcTipsVbaHereDocuments.html
'---------------------------------------------------------------------------------------

#Const HOST_IS_Access = False               ' �z�X�g�A�v���P�[�V����    : [True] Access, [False] ��Access
Private Const ENABLE_INSIDE_COMMENT = True  ' �q�A�h�L�������g���R�����g: [True] �L��, [False] ����
'#If HOST_IS_Access Then
'    Option Compare Database
'#End If

Option Explicit

Public Function HereDoc( _
    ByRef strIdentifier As String, _
    ByRef strModuleName As String, _
    ParamArray Params() As Variant _
    ) As String
    Dim sCode       As String
    Dim sFind       As String
    Dim nPos        As Long
    Dim sContent    As String
    Dim fStripLeadingSpaces As Boolean
    Dim avParams()  As Variant
    
'    With Application.VBE.ActiveVBProject.VBComponents(strModuleName).CodeModule
    With Workbooks("init.xla").VBProject.VBComponents(strModuleName).CodeModule
        sCode = vbCrLf & .Lines(1, .CountOfLines)
    End With

    avParams = Params

    ' �q�A�h�L�������g�p�̎��ʎq������(�\��1)
    ' -- '<<Identifier �܂��� '<<-Identifier �̎��s����
    sFind = "'<<"
    nPos = InStr(1, sCode, sFind, vbBinaryCompare)
    Do While nPos
        If HereDocMatchIdentifier(sCode, nPos, strIdentifier, fStripLeadingSpaces) Then
            sContent = HereDocGetContent(sCode, nPos, strIdentifier, fStripLeadingSpaces)
            sContent = HereDocFormatContent(sContent, avParams)
            HereDoc = sContent
            Exit Function
        End If
        nPos = InStr(nPos + Len(sFind), sCode, sFind, vbBinaryCompare)
    Loop

    ' �q�A�h�L�������g�p�̎��ʎq������(�\��2)
    ' -- HereDoc("Identifier" �̎��s����A�������R�����g�u���b�N
    sFind = "HereDoc(""" & strIdentifier & """"
    nPos = InStr(1, sCode, sFind, vbBinaryCompare)
    Do While nPos
        Select Case Mid$(sCode, nPos - 1, 1)
            Case " ", "(", "[", vbLf    ' �z���C�g�X�y�[�X��Z�p���[�^�ł���� OK
                If Not HereDocIsCommnetLine(sCode, nPos) Then
                    If Not HereDocIsStringLiteral(sCode, nPos) Then
                        Call HereDocMovePosToEndOfLogLine(sCode, nPos)
                        sContent = HereDocGetContent(sCode, nPos)
                        sContent = HereDocFormatContent(sContent, avParams)
                        HereDoc = sContent
                        Exit Function
                    End If
                End If
        End Select
        nPos = InStr(nPos + Len(sFind), sCode, sFind, vbBinaryCompare)
    Loop
End Function

Private Function HereDocMatchIdentifier( _
    sCode As String, _
    ByVal nPos As Long, _
    strIdentifier As String, _
    Optional ByRef StripLeadingSpaces As Boolean _
    ) As Boolean
    ' ���ʎq����v���邩
    ' -- ���݈ʒu������s��O�܂ł̕�������擾���A�g���~���O���Ĕ�r(�啶���E�����������)
    ' -- ���� StripLeadingSpaces �͏o�͗p
    ' -- ���� nPos �� ' ���w���Ă���Ƃ����O��
    Dim i As Long
    Dim S As String

    ' �s������ nPos �܂ł̂������ɃX�y�[�X�����ȊO���݂��Ȃ����Ƃ��m�F
    i = InStrRev(sCode, vbCrLf, nPos, vbBinaryCompare)
    If i = 0 Then   ' �擪�s�������ꍇ
        S = Left$(sCode, nPos - 1)
    Else
        S = Mid$(sCode, i + 2, nPos - (i + 2))
    End If
    If Trim$(S) <> vbNullString Then Exit Function

    nPos = nPos + 3
    StripLeadingSpaces = (Mid$(sCode, nPos, 1) = "-")
    If StripLeadingSpaces Then    ' �擪�X�y�[�X���폜����ꍇ(<<-)
        nPos = nPos + 1
    End If
    i = InStr(nPos, sCode, vbCrLf, vbBinaryCompare)
    S = Trim$(Mid$(sCode, nPos, i - nPos))
    
    HereDocMatchIdentifier = (StrComp(S, strIdentifier, vbBinaryCompare) = 0)
End Function

Private Function HereDocIsCommnetLine( _
    sCode As String, _
    nPos As Long _
    ) As Boolean
    Dim i As Long
    Dim S As String

    i = InStrRev(sCode, vbCrLf, nPos, vbBinaryCompare) + 2
    S = Mid$(sCode, i, nPos - i)
    HereDocIsCommnetLine = (Left$(LTrim$(S), 1) = "'")
End Function

Private Function HereDocIsStringLiteral( _
    sCode As String, _
    nPos As Long _
    ) As Boolean
    Dim i As Long
    Dim j As Long
    Dim S As String

    i = InStrRev(sCode, vbCrLf, nPos, vbBinaryCompare) + 2
    S = Mid$(sCode, i, nPos - i)

    i = InStr(1, S, """", vbBinaryCompare)
    Do While i
        j = j + 1
        i = InStr(i + 1, S, """", vbBinaryCompare)
    Loop

    HereDocIsStringLiteral = j \ 2
End Function

Private Sub HereDocMovePosToEndOfLogLine( _
    sCode As String, _
    nPos As Long _
    )
    Dim i As Long
    Dim S As String

    Do While True
        i = InStr(nPos, sCode, vbCrLf, vbBinaryCompare)
        S = Mid$(sCode, i - 2, 2)
        ' �����s�����s�p�������̏ꍇ
        If StrComp(S, " _", vbBinaryCompare) = 0 Then
            nPos = i + 2    ' ���s�擪�ֈړ�
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function HereDocGetContent( _
    sCode As String, _
    nPos As Long, _
    Optional CloseIdentifier As String, _
    Optional ByVal fStripLeadingSpaces As Boolean _
    ) As String
    ' ���e���擾
    ' -- �擾�J�n�ʒu ���� nPos �̎��s�擪
    ' -- �擾�I���ʒu ���� CloseIdentifier �ȗ����́A�A������R�����g�u���b�N�̏I�[
    '                 ���� CloseIdentifier �w�莞�́ACloseIdentifier �s�̑O�s�̏I�[
    ' -- ���� CloseIdentifier �w�莞�Ɍ���A���� fStripLeadingSpaces ���g�p����܂��B
    Dim i As Long
    Dim j As Long
    Dim l As Long
    Dim S As String
    Dim sFind As String
    Dim src As String
    Dim ret As String

    ' �擾�J�n�ʒu
    i = InStr(nPos, sCode, vbCrLf, vbBinaryCompare) + 2

    ' �A������R�����g�u���b�N�̏ꍇ
    If CloseIdentifier = vbNullString Then
        Do While True
            j = InStr(i, sCode, vbCrLf, vbBinaryCompare) + 2
            S = LTrim$(Mid$(sCode, i, j - i))
            If Left$(S, 1) = "'" Then
                If ENABLE_INSIDE_COMMENT Then   ' �q�A�h�L�������g���R�����g
                    If Mid$(S, 2, 1) <> "'" Then
                        ret = ret & Mid$(S, 2)
                    End If
                Else
                    ret = ret & Mid$(S, 2)
                End If
            Else
                Exit Do
            End If
            i = j
        Loop
    ' CloseIdentifier �s�̑O�s�܂ł̏ꍇ
    Else
        ' �I�[������
        sFind = "'" & CloseIdentifier & vbCrLf
        j = InStr(i, sCode, sFind, vbBinaryCompare)
        Do While j
            ' �s������ j �܂ł̂������ɃX�y�[�X�����ȊO���݂��Ȃ����Ƃ��m�F
            l = InStrRev(sCode, vbCrLf, j, vbBinaryCompare)
            S = Mid$(sCode, l + 2, j - (l + 2))
            If Trim$(S) = vbNullString Then Exit Do ' OK
            j = InStr(j + Len(sFind), sCode, sFind, vbBinaryCompare)
        Loop

        ' �I�[��������Ȃ��ꍇ�͎��s���G���[�𔭐�������
        If j = 0 Then
            Err.Raise 327, "HereDoc->HereDocGetContent", _
                      "�q�A�h�L�������g�̏I�[ '" & CloseIdentifier & "' ��������܂���B"
        End If

        ' �Y���͈͂̕�������擾
        src = Mid$(sCode, i, l + 2 - i) ' �㑱�̏����̓s����A�����ɉ��s���܂߂�

        ' 1 �s������
        i = 1
        l = Len(src)
        Do While i < l
            j = InStr(i, src, vbCrLf, vbBinaryCompare) + 2
            S = LTrim$(Mid$(src, i, j - i))
            If Left$(S, 1) = "'" Then
                S = Mid$(S, 2)
                If ENABLE_INSIDE_COMMENT Then   ' �q�A�h�L�������g���R�����g
                    If Left$(S, 1) = "'" Then
                        S = vbNullString
                    End If
                End If
            End If
            If fStripLeadingSpaces Then
                ret = ret & LTrim$(S)
            Else
                ret = ret & S
            End If
            i = j
        Loop
    End If

    ' �����̉��s�R�[�h���������ĕԋp
    If Len(ret) Then
        HereDocGetContent = Left$(ret, Len(ret) - 2)
    End If
End Function

Private Function HereDocFormatContent( _
    ByRef Source As String, _
    ByRef args() As Variant _
    ) As String
    ' ���e�������W�J���Đ��`
    Dim sContent As String

    ' "$$" �� Null �����ɒu�����܂��B
    sContent = replace(Source, "$$", vbNullChar, 1, -1, vbBinaryCompare)

    ' �����W�J
    If UBound(args) >= 0 Then
        sContent = ExpandVariables(sContent, args)
    End If

    #If HOST_IS_Access Then
        If InStr(1, sContent, "${", vbBinaryCompare) > 0 Then
            sContent = EvaluateExpression(sContent)
        End If
    #End If

    ' �c���� "$" ���폜���ANull ������ "$" �ɒu�����܂��B
    sContent = replace(sContent, "$", vbNullString, 1, -1, vbBinaryCompare)
    sContent = replace(sContent, vbNullChar, "$", 1, -1, vbBinaryCompare)

    HereDocFormatContent = sContent
End Function

Private Function ExpandVariables( _
    ByRef Source As String, _
    ByRef args() As Variant _
    ) As String
    ' �֐����F ExpandVariables
    ' ��  �I�F �C�ӌ��̈�����u�������������Ԃ��܂��B
    ' �쐬�ҁF YU-TANG@http://www.f3.dion.ne.jp/~element/msaccess/
    ' �쐬���F 2010/03/25
    ' �X�V�����F
    ' �߂�l�F String �^
    ' �����̐����F
    ' Source �u���Ώە�������w�肵�܂��B
    ' args  �u����̈�����C�ӌ��w�肵�܂��B
    '       �擪�̓v���[�X�z���_ $1 �ƒu������܂��B
    '       �Ȍ�A$2�A$3...�Ə��ɒu������܂��B
    ' �g�p��̒��ӁF
    '       Null ���n���ꂽ�ꍇ�� "Null" �ɒu������܂��B
    ' �g�p��F
    '  s = ExpandVariables("��$1-$2",Array("104","0123"))
    Dim i As Integer
    Dim S As String: S = Source

    ' �v���[�X�z���_ $n �������u�����܂��B
    ' -- $11 �� $1 �Œu�����Ȃ��悤�A��������u�����J�n���܂��B
    For i = UBound(args) To LBound(args) Step -1
        If IsNull(args(i)) Then
            S = replace(S, "$" & (i + 1), "Null", 1, -1, vbBinaryCompare)
        Else
            S = replace(S, "$" & (i + 1), args(i), 1, -1, vbBinaryCompare)
        End If
    Next

    ExpandVariables = S
End Function

#If HOST_IS_Access Then
    Private Function EvaluateExpression(ByRef Source As String) As String
        ' ����]�����܂��B�]����̒l�� Null �̏ꍇ�� "Null" �ɒu���������܂��B
        Dim i As Integer
        Dim j As Integer
        Dim S As String
        Dim exp As String
        Dim v As Variant
    
        S = Source
        i = InStr(1, S, "${", vbBinaryCompare)
    
        Do While i
            j = InStr(i + 2, S, "}", vbBinaryCompare)
            If j > 0 Then
                exp = Mid$(S, i + 2, j - (i + 2))
                v = Eval(exp)
                If IsNull(v) Then
                    v = "Null"
                End If
                S = Left$(S, i - 1) & v & Mid$(S, j + 1)
            Else
                Exit Do
            End If
            i = InStr(i, S, "${", vbBinaryCompare)
        Loop
    
        EvaluateExpression = S
    End Function
#End If
