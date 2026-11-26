Attribute VB_Name = "modConsolidador"
Option Explicit

'==========================================================================
'  CONSOLIDADOR DE RELATÓRIOS (Excel/VBA)  ·  Renan Vicente Braga
'--------------------------------------------------------------------------
'  Junta vários arquivos de extração (.xlsx / .xlsb / .xls / .csv) que
'  chegam com colunas em ordens diferentes e formatos bagunçados, e gera
'  um relatório único e padronizado, com conferência automática.
'
'  Tudo é configurado em abas, sem mexer no código:
'    PARAMETROS  pastas, linha do cabeçalho, nome do relatório
'    CONFIG      quais colunas o relatório tem e de onde cada uma vem
'    REGRAS      "de -> para" para padronizar textos (ex.: status)
'    BASE_*      tabelas de apoio para buscas (ex.: BASE_PRECOS)
'
'  Tipos aceitos na CONFIG (coluna TIPO):
'    TEXTO    copia como texto (preserva zeros à esquerda)
'    NUMERO   "000123" ou "1.234,50" -> número
'    DATA     "31/10/2026 18:48" ou número serial -> data sem hora
'    REGRA    aplica a aba REGRAS (sem regra = mantém o valor original)
'    BUSCA    PARAMETRO = ABA!colChave!colValor!COLUNA_DA_SAIDA_USADA_COMO_CHAVE
'    FIXO     PARAMETRO = valor fixo
'    ARQUIVO  nome do arquivo de origem (rastreabilidade)
'    FORMULA  PARAMETRO = fórmula com referências estruturadas, ex.:
'             =[@Quantidade]*[@[Valor unitário]]
'
'  Macros: Consolidar | Diagnosticar (não gera nada)
'==========================================================================

Private mRel As Collection        ' linhas do relatório de conferência
Private mAlerta As Boolean

Private Type ColCfg
    saida As String
    origem As String
    tipo As String
    param As String
End Type

Public Sub Consolidar()
    Executar False
End Sub

Public Sub Diagnosticar()
    Executar True
End Sub


Private Sub Executar(ByVal soDiag As Boolean)
    Dim pastaIn As String, pastaOut As String, linhaCab As Long, nomeRel As String
    Dim cfg() As ColCfg, nCol As Long, regras As Object, bases As Object
    Dim arquivos As Collection, f As Variant, t0 As Single
    Dim saida() As Variant, total As Long, r As Long, j As Long
    Dim blocos As New Collection, b As Variant, dados As Variant, mapa As Object
    Dim i As Long, k As Long, v As Variant, faltas As Object, semRegra As Object, invalidos As Object
    Dim wbOut As Workbook, ws As Worksheet, lo As ListObject, destino As String

    t0 = Timer
    Set mRel = New Collection: mAlerta = False
    Set faltas = CreateObject("Scripting.Dictionary")
    Set semRegra = CreateObject("Scripting.Dictionary")
    Set invalidos = CreateObject("Scripting.Dictionary")

    '--- parâmetros
    pastaIn = Caminho(Param("Pasta de entrada", "ENTRADAS"))
    pastaOut = Caminho(Param("Pasta de saída", "SAIDAS"))
    linhaCab = CLng(Val(Param("Linha do cabeçalho", "1")))
    nomeRel = Param("Nome do relatório", "RELATORIO CONSOLIDADO")
    If Dir(pastaOut, vbDirectory) = "" Then MkDir pastaOut

    cfg = LerConfig(nCol)
    Set regras = LerRegras()
    Set bases = CreateObject("Scripting.Dictionary")

    Set arquivos = ListarArquivos(pastaIn)
    If arquivos.Count = 0 Then
        MsgBox "Nenhum arquivo em " & pastaIn, vbExclamation, "Consolidador"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Falha

    '--- 1) leitura de todos os arquivos
    For Each f In arquivos
        Application.StatusBar = "Lendo " & NomeArq(CStr(f)) & "..."
        dados = LerArquivo(CStr(f))
        If IsArray(dados) Then
            Set mapa = MapaCabecalho(dados, linhaCab)
            k = UBound(dados, 1) - linhaCab
            ' confere se as colunas pedidas na CONFIG existem neste arquivo
            For j = 1 To nCol
                If NecessitaOrigem(cfg(j).tipo) Then
                    If Not mapa.Exists(Normalizar(cfg(j).origem)) Then _
                        Check "ATENÇÃO", NomeArq(CStr(f)), "coluna [" & cfg(j).origem & "] não encontrada (fica vazia)"
                End If
            Next
            blocos.Add Array(dados, mapa, NomeArq(CStr(f)), k)
            total = total + k
            Check "OK", NomeArq(CStr(f)), Format(k, "#,##0") & " linhas lidas"
        End If
    Next
    If total = 0 Then Err.Raise vbObjectError + 1, , "Os arquivos não têm linhas de dados."

    '--- 2) monta a tabela padronizada na memória
    ReDim saida(1 To total, 1 To nCol)
    r = 0
    For Each b In blocos
        dados = b(0): Set mapa = b(1)
        For i = linhaCab + 1 To UBound(dados, 1)
            If LinhaVazia(dados, i) Then GoTo Proxima
            r = r + 1
            For j = 1 To nCol
                v = Empty
                If NecessitaOrigem(cfg(j).tipo) Then
                    If mapa.Exists(Normalizar(cfg(j).origem)) Then v = dados(i, mapa(Normalizar(cfg(j).origem)))
                End If
                Select Case cfg(j).tipo
                    Case "TEXTO": If Not IsError(v) Then v = Trim$(Texto(v))
                    Case "NUMERO"
                        v = ParaNumero(v)
                        If VarType(v) = vbString And Trim$(v) <> "" Then invalidos(cfg(j).saida) = invalidos(cfg(j).saida) + 1
                    Case "DATA"
                        v = ParaData(v)
                        If VarType(v) = vbString And Trim$(v) <> "" Then invalidos(cfg(j).saida) = invalidos(cfg(j).saida) + 1
                    Case "REGRA"
                        If regras.Exists(Normalizar(v)) Then
                            v = regras(Normalizar(v))
                        ElseIf Trim$(Texto(v)) <> "" Then
                            semRegra(Trim$(Texto(v))) = semRegra(Trim$(Texto(v))) + 1
                        End If
                    Case "FIXO": v = cfg(j).param
                    Case "ARQUIVO": v = b(2)
                    Case "BUSCA", "FORMULA": v = Empty          ' resolvidas depois
                End Select
                saida(r, j) = v
            Next
Proxima:
        Next
    Next
    total = r

    '--- 3) buscas (dependem de outras colunas da saída)
    For j = 1 To nCol
        If cfg(j).tipo = "BUSCA" Then ResolverBusca cfg, nCol, j, saida, total, faltas
    Next

    '--- 4) relatório de conferência
    Dim c As Variant
    For Each c In invalidos.Keys
        Check "ATENÇÃO", "Coluna " & c, invalidos(c) & " valor(es) que não puderam ser convertidos (mantidos como texto)"
    Next
    For Each c In faltas.Keys
        Check "ATENÇÃO", "Busca " & c, Resumo(faltas(c), "chave(s) sem cadastro na base")
    Next
    If semRegra.Count > 0 Then Check "INFO", "Padronização", Resumo(semRegra, "valor(es) sem regra na aba REGRAS (mantidos como vieram)")

    If soDiag Then
        Check "INFO", "Diagnóstico", Format(total, "#,##0") & " linhas seriam geradas. Nada foi salvo."
        GoTo Fim
    End If

    '--- 5) grava o relatório num arquivo novo, com data e hora no nome
    Application.StatusBar = "Gerando o relatório..."
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set ws = wbOut.Worksheets(1): ws.Name = "DADOS"
    For j = 1 To nCol
        ws.Cells(1, j).Value = cfg(j).saida
        Select Case cfg(j).tipo
            Case "TEXTO": ws.Columns(j).NumberFormat = "@"
            Case "DATA": ws.Columns(j).NumberFormat = "dd/mm/yyyy"
            Case "NUMERO", "BUSCA": ws.Columns(j).NumberFormat = "#,##0.00"
        End Select
    Next
    ws.Range("A2").Resize(total, nCol).Value = saida
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").Resize(total + 1, nCol), , xlYes)
    lo.Name = "tbDados": lo.TableStyle = "TableStyleMedium2"
    For j = 1 To nCol
        If cfg(j).tipo = "FORMULA" Then
            lo.ListColumns(j).DataBodyRange.Formula = cfg(j).param
            lo.ListColumns(j).DataBodyRange.NumberFormat = "#,##0.00"
        End If
    Next
    ws.Columns.AutoFit
    CriarResumo wbOut, cfg, nCol
    Application.Calculate

    destino = pastaOut & nomeRel & " - " & Format(Now, "dd_mm_yyyy") & " - " & Format(Now, "hh""h""nn") & ".xlsx"
    wbOut.SaveAs Filename:=destino, FileFormat:=xlOpenXMLWorkbook
    Check "OK", "Relatório gerado", NomeArq(destino) & "  (" & Format(total, "#,##0") & " linhas)"

Fim:
    Restaurar
    EscreverResultado soDiag, total, Timer - t0
    If Not soDiag Then
        wbOut.Activate
        MsgBox "Relatório gerado:" & vbLf & NomeArq(destino) & vbLf & vbLf & _
               IIf(mAlerta, "Há pontos para conferir na aba RESULTADO.", "Conferência OK."), _
               IIf(mAlerta, vbExclamation, vbInformation), "Consolidador"
    End If
    Exit Sub

Falha:
    Dim msg As String: msg = Err.Description
    Restaurar
    Check "ERRO", "Execução interrompida", msg
    EscreverResultado soDiag, total, Timer - t0
    MsgBox "Erro: " & msg, vbCritical, "Consolidador"
End Sub


'============================ CONFIGURAÇÃO ============================

Private Function Param(ByVal nome As String, ByVal padrao As String) As String
    Dim ws As Worksheet, achou As Range
    Param = padrao
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PARAMETROS")
    Set achou = ws.Columns(1).Find(What:=nome, LookAt:=xlWhole)
    If Not achou Is Nothing Then
        If Trim$(CStr(achou.Offset(0, 1).Value)) <> "" Then Param = Trim$(CStr(achou.Offset(0, 1).Value))
    End If
End Function

' caminho relativo -> pasta deste arquivo
Private Function Caminho(ByVal p As String) As String
    If InStr(p, ":\") = 0 And Left$(p, 2) <> "\\" Then p = ThisWorkbook.Path & "\" & p
    If Right$(p, 1) <> "\" Then p = p & "\"
    Caminho = p
End Function

Private Function LerConfig(ByRef n As Long) As ColCfg()
    Dim ws As Worksheet, ult As Long, i As Long, arr() As ColCfg
    Set ws = ThisWorkbook.Worksheets("CONFIG")
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then Err.Raise vbObjectError + 2, , "A aba CONFIG está vazia."
    ReDim arr(1 To ult - 1)
    For i = 2 To ult
        If Trim$(CStr(ws.Cells(i, 1).Value)) <> "" Then
            n = n + 1
            arr(n).saida = Trim$(CStr(ws.Cells(i, 1).Value))
            arr(n).origem = Trim$(CStr(ws.Cells(i, 2).Value))
            arr(n).tipo = UCase$(Trim$(CStr(ws.Cells(i, 3).Value)))
            arr(n).param = Trim$(CStr(ws.Cells(i, 4).Value))
            If arr(n).tipo = "" Then arr(n).tipo = "TEXTO"
        End If
    Next
    ReDim Preserve arr(1 To n)
    LerConfig = arr
End Function

Private Function NecessitaOrigem(ByVal tipo As String) As Boolean
    NecessitaOrigem = (tipo = "TEXTO" Or tipo = "NUMERO" Or tipo = "DATA" Or tipo = "REGRA")
End Function

Private Function LerRegras() As Object
    Dim d As Object, ws As Worksheet, ult As Long, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("REGRAS")
    On Error GoTo 0
    If ws Is Nothing Then Set LerRegras = d: Exit Function
    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ult
        If Trim$(CStr(ws.Cells(i, 1).Value)) <> "" Then d(Normalizar(ws.Cells(i, 1).Value)) = Trim$(CStr(ws.Cells(i, 2).Value))
    Next
    Set LerRegras = d
End Function

' BUSCA: PARAMETRO = ABA!colChave!colValor!COLUNA_DA_SAIDA
Private Sub ResolverBusca(cfg() As ColCfg, ByVal nCol As Long, ByVal j As Long, saida() As Variant, _
                          ByVal total As Long, faltas As Object)
    Dim p As Variant, ws As Worksheet, base As Object, dados As Variant, i As Long, ult As Long
    Dim colChave As Long, colValor As Long, jChave As Long, k As String, d As Object
    p = Split(cfg(j).param, "!")
    If UBound(p) < 3 Then Err.Raise vbObjectError + 3, , "BUSCA de [" & cfg(j).saida & "]: use ABA!colChave!colValor!COLUNA"
    Set ws = ThisWorkbook.Worksheets(CStr(p(0)))
    colChave = CLng(p(1)): colValor = CLng(p(2))
    For i = 1 To nCol
        If Normalizar(cfg(i).saida) = Normalizar(p(3)) Then jChave = i
    Next
    If jChave = 0 Then Err.Raise vbObjectError + 4, , "BUSCA de [" & cfg(j).saida & "]: coluna-chave [" & p(3) & "] não existe na CONFIG."
    Set base = CreateObject("Scripting.Dictionary")
    ult = ws.Cells(ws.Rows.Count, colChave).End(xlUp).Row
    If ult >= 2 Then
        dados = ws.Range(ws.Cells(1, 1), ws.Cells(ult, IIf(colValor > colChave, colValor, colChave))).Value
        For i = 2 To ult
            k = Chave(dados(i, colChave))
            If k <> "" And Not base.Exists(k) Then base(k) = dados(i, colValor)
        Next
    End If
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To total
        k = Chave(saida(i, jChave))
        If base.Exists(k) Then
            saida(i, j) = base(k)
        Else
            d(k) = d(k) + 1
        End If
    Next
    If d.Count > 0 Then Set faltas(cfg(j).saida) = d
End Sub


'============================ LEITURA ============================

Private Function ListarArquivos(ByVal pasta As String) As Collection
    Dim c As New Collection, nome As String
    nome = Dir(pasta & "*.*")
    Do While nome <> ""
        If Left$(nome, 2) <> "~$" And (LCase$(nome) Like "*.xls*" Or LCase$(nome) Like "*.csv") Then c.Add pasta & nome
        nome = Dir()
    Loop
    Set ListarArquivos = c
End Function

Private Function LerArquivo(ByVal caminho As String) As Variant
    Dim wb As Workbook
    Set wb = Workbooks.Open(Filename:=caminho, UpdateLinks:=0, ReadOnly:=True, Local:=True)
    LerArquivo = wb.Worksheets(1).UsedRange.Value
    wb.Close SaveChanges:=False
End Function

Private Function MapaCabecalho(dados As Variant, ByVal lh As Long) As Object
    Dim m As Object, j As Long, h As String
    Set m = CreateObject("Scripting.Dictionary")
    For j = 1 To UBound(dados, 2)
        h = Normalizar(dados(lh, j))
        If h <> "" And Not m.Exists(h) Then m(h) = j
    Next
    Set MapaCabecalho = m
End Function

Private Function LinhaVazia(dados As Variant, ByVal i As Long) As Boolean
    Dim j As Long
    For j = 1 To UBound(dados, 2)
        If Not IsEmpty(dados(i, j)) Then
            If Trim$(Texto(dados(i, j))) <> "" Then Exit Function
        End If
    Next
    LinhaVazia = True
End Function


'============================ CONVERSÕES ============================

Private Function Texto(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbDouble Then
        If v = Int(v) And Abs(v) < 1E+15 Then Texto = Format(v, "0"): Exit Function
    End If
    Texto = CStr(v)
End Function

' "000123" -> 123 | "1.234,50" -> 1234,5 | texto inválido fica como veio
Private Function ParaNumero(ByVal v As Variant) As Variant
    Dim s As String
    ParaNumero = v
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) <> vbString Then Exit Function
    s = Trim$(Replace(Replace(v, "R$", ""), " ", ""))
    If s = "" Then ParaNumero = Empty: Exit Function
    If InStr(s, ",") > 0 Then s = Replace(Replace(s, ".", ""), ",", ".")
    If s Like "*[!0-9.-]*" Then Exit Function
    ParaNumero = Val(s)
End Function

' "31/10/2026 18:48:46" -> 31/10/2026 | "5/1/2026" -> 05/01/2026 | serial -> data | AAAA-MM-DD -> data
Private Function ParaData(ByVal v As Variant) As Variant
    Dim s As String, p As Variant
    ParaData = v
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then ParaData = DateSerial(Year(v), Month(v), Day(v)): Exit Function
    If VarType(v) <> vbString Then
        If IsNumeric(v) Then ParaData = CDate(Int(v))
        Exit Function
    End If
    s = Split(Trim$(v), " ")(0)
    p = Split(s, "/")
    If UBound(p) = 2 Then
        If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) Then ParaData = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))
    ElseIf s Like "####-##-##*" Then
        ParaData = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
    End If
End Function

' chave de busca: maiúsculas, sem espaços e sem zeros à esquerda em códigos numéricos
Private Function Chave(ByVal v As Variant) As String
    Dim s As String
    s = UCase$(Trim$(Texto(v)))
    If Len(s) > 0 And Not s Like "*[!0-9]*" Then
        Do While Len(s) > 1 And Left$(s, 1) = "0": s = Mid$(s, 2): Loop
    End If
    Chave = s
End Function

' maiúsculas, sem acentos, sem espaços duplicados (para casar cabeçalhos e regras)
Private Function Normalizar(ByVal v As Variant) As String
    Dim s As String, com As String, i As Long
    Const SEM As String = "AAAAAEEEEIIIIOOOOOUUUUC"
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    s = UCase$(Trim$(CStr(v)))
    com = ChrW(193) & ChrW(192) & ChrW(194) & ChrW(195) & ChrW(196) & ChrW(201) & ChrW(200) & ChrW(202) & ChrW(203) & _
          ChrW(205) & ChrW(204) & ChrW(206) & ChrW(207) & ChrW(211) & ChrW(210) & ChrW(212) & ChrW(213) & ChrW(214) & _
          ChrW(218) & ChrW(217) & ChrW(219) & ChrW(220) & ChrW(199)
    For i = 1 To Len(com): s = Replace(s, Mid$(com, i, 1), Mid$(SEM, i, 1)): Next
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    Normalizar = s
End Function


'============================ SAÍDA ============================

' aba RESUMO no relatório: total por cada coluna do tipo REGRA (ex.: status)
Private Sub CriarResumo(wb As Workbook, cfg() As ColCfg, ByVal nCol As Long)
    Dim ws As Worksheet, j As Long, r As Long, d As Object, lo As ListObject, v As Variant, c As Variant, jVal As Long
    Set lo = wb.Worksheets("DADOS").ListObjects(1)
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): ws.Name = "RESUMO"
    For j = 1 To nCol
        If cfg(j).tipo = "FORMULA" Or cfg(j).tipo = "BUSCA" Then jVal = j
    Next
    r = 1
    For j = 1 To nCol
        If cfg(j).tipo = "REGRA" Then
            ws.Cells(r, 1).Value = cfg(j).saida: ws.Cells(r, 2).Value = "Linhas"
            If jVal > 0 Then ws.Cells(r, 3).Value = cfg(jVal).saida
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Bold = True
            Set d = CreateObject("Scripting.Dictionary")
            If lo.ListRows.Count = 1 Then
                d(CStr(lo.ListColumns(j).DataBodyRange.Value)) = 1
            Else
                For Each v In lo.ListColumns(j).DataBodyRange.Value
                    d(CStr(v)) = 1
                Next
            End If
            For Each c In d.Keys
                r = r + 1
                ws.Cells(r, 1).Value = c
                ws.Cells(r, 2).Formula = "=COUNTIFS(tbDados[" & cfg(j).saida & "]," & ws.Cells(r, 1).Address(False, False) & ")"
                If jVal > 0 Then ws.Cells(r, 3).Formula = "=SUMIFS(tbDados[" & cfg(jVal).saida & "],tbDados[" & cfg(j).saida & "]," & ws.Cells(r, 1).Address(False, False) & ")"
            Next
            r = r + 2
        End If
    Next
    ws.Columns("C").NumberFormat = "#,##0.00"
    ws.Columns("A:C").AutoFit
End Sub

Private Sub Check(ByVal status As String, ByVal item As String, ByVal detalhe As String)
    mRel.Add Array(status, item, detalhe)
    If status = "ATENÇÃO" Or status = "ERRO" Then mAlerta = True
End Sub

Private Function Resumo(d As Object, ByVal txt As String) As String
    Dim c As Variant, n As Long, ex As String
    For Each c In d.Keys
        If n < 8 Then ex = ex & IIf(ex = "", "", ", ") & c
        n = n + 1
    Next
    Resumo = n & " " & txt & ": " & ex & IIf(n > 8, "...", "")
End Function

Private Sub EscreverResultado(ByVal soDiag As Boolean, ByVal total As Long, ByVal seg As Single)
    Dim ws As Worksheet, r As Long, c As Variant, cor As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("RESULTADO")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "RESULTADO"
    End If
    ws.Cells.Clear
    ws.Range("A1").Value = IIf(soDiag, "DIAGNÓSTICO", "CONSOLIDAÇÃO") & "  ·  " & Format(Now, "dd/mm/yyyy hh:nn") & _
                           "  ·  " & Format(total, "#,##0") & " linhas  ·  " & Format(seg, "0.0") & " s"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Range("A3:C3").Value = Array("STATUS", "ITEM", "DETALHE")
    ws.Range("A3:C3").Font.Bold = True
    r = 3
    For Each c In mRel
        r = r + 1
        ws.Cells(r, 1).Value = c(0): ws.Cells(r, 2).Value = c(1): ws.Cells(r, 3).Value = c(2)
        Select Case c(0)
            Case "OK": cor = RGB(226, 239, 218)
            Case "ATENÇÃO": cor = RGB(255, 242, 204)
            Case "ERRO": cor = RGB(248, 215, 218)
            Case Else: cor = RGB(242, 242, 242)
        End Select
        ws.Cells(r, 1).Interior.Color = cor
        ws.Cells(r, 1).Font.Bold = True
    Next
    ws.Columns("A").ColumnWidth = 12: ws.Columns("B").ColumnWidth = 30: ws.Columns("C").ColumnWidth = 100
    ws.Activate
End Sub

Private Sub Restaurar()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
End Sub

Private Function NomeArq(ByVal p As String) As String
    NomeArq = Mid$(p, InStrRev(p, "\") + 1)
End Function
