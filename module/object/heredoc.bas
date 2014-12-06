Attribute VB_Name = "object_heredoc"
'---------------------------------------------------------------------------------------
' Module    : modHereDoc
' Version   : 1.0.0
' DateTime  : 2010/03/27 10:41
' Author    : YU-TANG
' Purpose   : VBA 用ヒアドキュメント
' Reference : http://www.f3.dion.ne.jp/~element/msaccess/AcTipsVbaHereDocuments.html
'---------------------------------------------------------------------------------------

#Const HOST_IS_Access = False               ' ホストアプリケーション    : [True] Access, [False] 非Access
Private Const ENABLE_INSIDE_COMMENT = True  ' ヒアドキュメント内コメント: [True] 有効, [False] 無効
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

    ' ヒアドキュメント用の識別子を検索(構文1)
    ' -- '<<Identifier または '<<-Identifier の次行から
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

    ' ヒアドキュメント用の識別子を検索(構文2)
    ' -- HereDoc("Identifier" の次行から連続したコメントブロック
    sFind = "HereDoc(""" & strIdentifier & """"
    nPos = InStr(1, sCode, sFind, vbBinaryCompare)
    Do While nPos
        Select Case Mid$(sCode, nPos - 1, 1)
            Case " ", "(", "[", vbLf    ' ホワイトスペースやセパレータであれば OK
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
    ' 識別子が一致するか
    ' -- 現在位置から改行手前までの文字列を取得し、トリミングして比較(大文字・小文字を区別)
    ' -- 引数 StripLeadingSpaces は出力用
    ' -- 引数 nPos は ' を指しているという前提
    Dim i As Long
    Dim S As String

    ' 行頭から nPos までのあいだにスペース文字以外存在しないことを確認
    i = InStrRev(sCode, vbCrLf, nPos, vbBinaryCompare)
    If i = 0 Then   ' 先頭行だった場合
        S = Left$(sCode, nPos - 1)
    Else
        S = Mid$(sCode, i + 2, nPos - (i + 2))
    End If
    If Trim$(S) <> vbNullString Then Exit Function

    nPos = nPos + 3
    StripLeadingSpaces = (Mid$(sCode, nPos, 1) = "-")
    If StripLeadingSpaces Then    ' 先頭スペースを削除する場合(<<-)
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
        ' 物理行末が行継続文字の場合
        If StrComp(S, " _", vbBinaryCompare) = 0 Then
            nPos = i + 2    ' 次行先頭へ移動
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
    ' 内容を取得
    ' -- 取得開始位置 引数 nPos の次行先頭
    ' -- 取得終了位置 引数 CloseIdentifier 省略時は、連続するコメントブロックの終端
    '                 引数 CloseIdentifier 指定時は、CloseIdentifier 行の前行の終端
    ' -- 引数 CloseIdentifier 指定時に限り、引数 fStripLeadingSpaces が使用されます。
    Dim i As Long
    Dim j As Long
    Dim l As Long
    Dim S As String
    Dim sFind As String
    Dim src As String
    Dim ret As String

    ' 取得開始位置
    i = InStr(nPos, sCode, vbCrLf, vbBinaryCompare) + 2

    ' 連続するコメントブロックの場合
    If CloseIdentifier = vbNullString Then
        Do While True
            j = InStr(i, sCode, vbCrLf, vbBinaryCompare) + 2
            S = LTrim$(Mid$(sCode, i, j - i))
            If Left$(S, 1) = "'" Then
                If ENABLE_INSIDE_COMMENT Then   ' ヒアドキュメント内コメント
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
    ' CloseIdentifier 行の前行までの場合
    Else
        ' 終端を検索
        sFind = "'" & CloseIdentifier & vbCrLf
        j = InStr(i, sCode, sFind, vbBinaryCompare)
        Do While j
            ' 行頭から j までのあいだにスペース文字以外存在しないことを確認
            l = InStrRev(sCode, vbCrLf, j, vbBinaryCompare)
            S = Mid$(sCode, l + 2, j - (l + 2))
            If Trim$(S) = vbNullString Then Exit Do ' OK
            j = InStr(j + Len(sFind), sCode, sFind, vbBinaryCompare)
        Loop

        ' 終端が見つからない場合は実行時エラーを発生させる
        If j = 0 Then
            Err.Raise 327, "HereDoc->HereDocGetContent", _
                      "ヒアドキュメントの終端 '" & CloseIdentifier & "' が見つかりません。"
        End If

        ' 該当範囲の文字列を取得
        src = Mid$(sCode, i, l + 2 - i) ' 後続の処理の都合上、末尾に改行を含める

        ' 1 行ずつ処理
        i = 1
        l = Len(src)
        Do While i < l
            j = InStr(i, src, vbCrLf, vbBinaryCompare) + 2
            S = LTrim$(Mid$(src, i, j - i))
            If Left$(S, 1) = "'" Then
                S = Mid$(S, 2)
                If ENABLE_INSIDE_COMMENT Then   ' ヒアドキュメント内コメント
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

    ' 末尾の改行コードを除去して返却
    If Len(ret) Then
        HereDocGetContent = Left$(ret, Len(ret) - 2)
    End If
End Function

Private Function HereDocFormatContent( _
    ByRef Source As String, _
    ByRef args() As Variant _
    ) As String
    ' 内容を引数展開して整形
    Dim sContent As String

    ' "$$" を Null 文字に置換します。
    sContent = replace(Source, "$$", vbNullChar, 1, -1, vbBinaryCompare)

    ' 引数展開
    If UBound(args) >= 0 Then
        sContent = ExpandVariables(sContent, args)
    End If

    #If HOST_IS_Access Then
        If InStr(1, sContent, "${", vbBinaryCompare) > 0 Then
            sContent = EvaluateExpression(sContent)
        End If
    #End If

    ' 残った "$" を削除し、Null 文字を "$" に置換します。
    sContent = replace(sContent, "$", vbNullString, 1, -1, vbBinaryCompare)
    sContent = replace(sContent, vbNullChar, "$", 1, -1, vbBinaryCompare)

    HereDocFormatContent = sContent
End Function

Private Function ExpandVariables( _
    ByRef Source As String, _
    ByRef args() As Variant _
    ) As String
    ' 関数名： ExpandVariables
    ' 目  的： 任意個数の引数を置換した文字列を返します。
    ' 作成者： YU-TANG@http://www.f3.dion.ne.jp/~element/msaccess/
    ' 作成日： 2010/03/25
    ' 更新履歴：
    ' 戻り値： String 型
    ' 引数の説明：
    ' Source 置換対象文字列を指定します。
    ' args  置換後の引数を任意個数指定します。
    '       先頭はプレースホルダ $1 と置換されます。
    '       以後、$2、$3...と順に置換されます。
    ' 使用上の注意：
    '       Null が渡された場合は "Null" に置換されます。
    ' 使用例：
    '  s = ExpandVariables("〒$1-$2",Array("104","0123"))
    Dim i As Integer
    Dim S As String: S = Source

    ' プレースホルダ $n を順次置換します。
    ' -- $11 を $1 で置換しないよう、末尾から置換を開始します。
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
        ' 式を評価します。評価後の値が Null の場合は "Null" に置き換えられます。
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
