# ============================================================
# Dashboard Nintendo - PDV / Estoque / Ruptura / Display / Cobertura
# Consulta o bi_cliente e gera um bloco JS (const DADOS = {...})
# injetado dentro do dashboard.html (protótipo standalone, sem servidor).
#
# Uso:
#   .\atualizar_dashboard.ps1
#
# Fonte dos indicadores:
#   bi_f_produto      -> presenca (FL_PRESENTE), ruptura (FL_RUPTURA),
#                         frentes (FRENTES) por loja/SKU, coletado a
#                         cada visita do promotor. ESTOQUE_VIRTUAL nao
#                         esta populado nesta base (fica sempre 0), por
#                         isso NAO é usado como quantidade de estoque.
#   bi_f_formulario   -> formulario 3 (MOBILIARIO), pergunta 9
#                         "LOJA TEM MOBILIARIO DE GAMES?" = proxy de
#                         display/material ativo na loja.
#   bi_f_efetividade  -> planejado x realizado (roteiro) por loja/semana.
#   bi_d_tempo        -> calendario (semana atual, semana anterior).
#
# IMPORTANTE - "mix pertinente" da ruptura:
#   A base testa TODOS os ~104 SKUs do catalogo em toda visita, mesmo
#   em lojas que nunca venderiam boa parte deles. Ruptura "crua"
#   (FL_RUPTURA sobre todos os SKUs) fica ~95% e não serve como KPI.
#   Por isso calculamos ruptura só sobre o "mix pertinente" de cada
#   loja = SKUs que a própria loja já teve PRESENTE ao menos 1x nos
#   últimos 90 dias. É uma aproximação (não é o mix oficial da rede,
#   que ainda depende do de-para que o comercial/sede vai validar) mas
#   já tira o ruído dos SKUs que a loja nunca carrega.
#
# FILTRO DE MÊS:
#   O dashboard também guarda snapshots de meses anteriores, cada um
#   recalculado como se "hoje" fosse o último dia daquele mês (mesma
#   janela de mix/visita, só que ancorada numa data passada). O roteiro
#   (planejado x realizado) também é por mês: soma quantas semanas de
#   bi_f_efetividade dentro daquele mês foram planejadas/realizadas -
#   "atual" soma do 1o dia do mês corrente até hoje (mês em andamento).
#
# NÃO incluído (não existe nesta base, precisa do arquivo semanal do
# cliente / dados de sell-out):
#   - Estoque em unidades por SKU (loja a loja)
#   - Vendas semanais / giro / cobertura de estoque em dias
#   - Excesso de estoque x baixo giro
# ============================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$htmlFile = Join-Path $root 'dashboard.html'
$janelaMixDias = 90
$janelaVisitaDias = 60
$numMesesHistorico = 3

Write-Host "Dashboard Nintendo - atualizando indicadores" -ForegroundColor Magenta

$cs = "Server=SEU_SERVIDOR_SQL;Database=bi_cliente;Integrated Security=SSPI;TrustServerCertificate=True;Connection Timeout=30"
$cn = New-Object System.Data.SqlClient.SqlConnection $cs
$cn.Open()

function Exec([string]$sql) {
    $cmd = $cn.CreateCommand(); $cmd.CommandTimeout = 300; $cmd.CommandText = $sql
    [void]$cmd.ExecuteNonQuery()
}
function Invoke-Sql([string]$sql) {
    $cmd = $cn.CreateCommand(); $cmd.CommandTimeout = 300; $cmd.CommandText = $sql
    $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $dt = New-Object System.Data.DataTable
    [void]$da.Fill($dt)
    return , $dt
}
function S([object]$v) { if ($null -eq $v -or $v -is [dbnull]) { return '' }; return "$v".Trim() }
function N([object]$v) { if ($null -eq $v -or $v -is [dbnull]) { return 0 }; return $v }
function Dt([datetime]$d) { return $d.ToString('yyyy-MM-dd') }

# links das fotos (pergunta tipo "photo") tiradas na ultima resposta do formulario em cada loja, na janela
# [ini, ref). O servidor de imagens redireciona http -> https; o site e https, entao ja grava em https.
function FotosUltimaVisita([int]$form, [int]$pergChave, [int]$pergFoto, [string]$ini, [string]$ref) {
    $rows = Invoke-Sql @"
;WITH ult AS (
  SELECT COD_LOJA, MAX(DATA) AS ultima_data
  FROM bi_f_formulario WHERE COD_FORMULARIO = $form AND COD_PERGUNTA = $pergChave AND DATA >= '$ini' AND DATA < '$ref'
  GROUP BY COD_LOJA
)
SELECT f.COD_LOJA, r.DES_RESPOSTA url
FROM ult u
JOIN bi_f_formulario f ON f.COD_LOJA = u.COD_LOJA AND f.DATA = u.ultima_data AND f.COD_FORMULARIO = $form AND f.COD_PERGUNTA = $pergFoto
JOIN bi_d_formulario_respostas r ON r.bi_id_formulario_respostas = f.bi_id_formulario_respostas
WHERE r.DES_RESPOSTA LIKE 'http%'
"@
    $mapa = @{}
    foreach ($r in $rows.Rows) {
        $loja = S $r.COD_LOJA
        $url = (S $r.url) -replace '^http://', 'https://'
        if (-not $mapa.ContainsKey($loja)) { $mapa[$loja] = New-Object System.Collections.ArrayList }
        # no maximo 6 por loja pra nao inchar os docs do Firestore
        if ($mapa[$loja].Count -lt 6 -and -not $mapa[$loja].Contains($url)) { [void]$mapa[$loja].Add($url) }
    }
    return $mapa
}

# apelidos de exibicao para nomes de rede estranhos no bi_cliente (nao mexe no banco do cliente, so no rotulo mostrado)
$APELIDOS_REDE = @{ 'TIM ALEGRIA' = 'TIM' }
function NomeRede([string]$v) { if ($APELIDOS_REDE.ContainsKey($v)) { return $APELIDOS_REDE[$v] }; return $v }

# ---------- lista de meses (mais recente primeiro; "atual" sempre em 1o lugar) ----------
$MESES_PT = @('JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ')
$primeiroDiaMesAtual = (Get-Date).Date.AddDays( -1 * ((Get-Date).Day - 1) )
$mesesConfig = New-Object System.Collections.ArrayList
# "ref" = limite superior EXCLUSIVO da janela (tudo que aconteceu antes de "ref" conta).
# "inicio" = primeiro dia do mes, usado pra somar as semanas de roteiro daquele mes.
# pro mes atual usamos amanha como ref, pra garantir que o dia de hoje entra na janela
# (ou seja, "atual" = mes corrente ate hoje, nao um mes fechado).
[void]$mesesConfig.Add([ordered]@{ chave = 'atual'; rotulo = 'Atual'; ehAtual = $true; inicio = $primeiroDiaMesAtual; ref = (Get-Date).Date.AddDays(1) })
for ($i = 1; $i -le $numMesesHistorico; $i++) {
    $inicioMes = $primeiroDiaMesAtual.AddMonths(-$i)
    $refExclusivo = $inicioMes.AddMonths(1)
    $chave = $inicioMes.ToString('yyyy-MM')
    $rotulo = "$($MESES_PT[$inicioMes.Month-1])/$($inicioMes.ToString('yy'))"
    [void]$mesesConfig.Add([ordered]@{ chave = $chave; rotulo = $rotulo; ehAtual = $false; inicio = $inicioMes; ref = $refExclusivo })
}

# ---------- 6) catalogo de SKUs (para o filtro) ----------
Write-Host "  consultando catalogo de SKUs..."
$skusCat = Invoke-Sql @"
SELECT DES_SKU FROM bi_d_sku WHERE bi_id_sku > 3 ORDER BY DES_SKU
"@
$skus = @($skusCat.Rows | ForEach-Object { S $_.DES_SKU } | Where-Object { $_ })

# ---------- 7) campanhas ativas agora (Visao Geral) ----------
# cada campanha vira um formulario separado no bi_cliente (ex: "CAMPANHA: DIAS DOS PAIS"),
# com a pergunta 51 "LOJA TEM IMPLEMENTACAO DE CAMPANHA?" e a 100 pro motivo de nao ter
# ("MOTIVO DE NAO IMPLEMENTACAO" - o form antigo usava a pergunta 24 pra isso, mas as
# campanhas atuais (59/60/61) usam a 100; 24 nao tem nenhuma resposta nelas).
# consideramos "ativa" quem ainda esta sendo respondida nos ultimos 14 dias - isso e sempre
# ao vivo (nao existe "campanha retroativa" por mes, igual cobertura/roteiro).
Write-Host "  consultando campanhas ativas..."
$campanhaRespostas = Invoke-Sql @"
;WITH ativos AS (
  SELECT COD_FORMULARIO
  FROM bi_f_formulario
  WHERE COD_PERGUNTA = 51
  GROUP BY COD_FORMULARIO
  HAVING MAX(DATA) >= CONVERT(varchar, DATEADD(day,-14,GETDATE()), 23)
),
ultimaPorLoja AS (
  -- so conta loja aqui se ela de fato respondeu nos ultimos 14 dias - sem esse filtro,
  -- uma resposta de meses atras (de qualquer loja do formulario) contava como "ativa" pra
  -- sempre, so por causa do form como um todo ter tido alguma atividade recente em OUTRA loja.
  SELECT f.COD_FORMULARIO, f.COD_LOJA, MAX(f.DATA) ultima_data
  FROM bi_f_formulario f
  JOIN ativos at ON at.COD_FORMULARIO = f.COD_FORMULARIO
  WHERE f.COD_PERGUNTA = 51 AND f.DATA >= CONVERT(varchar, DATEADD(day,-14,GETDATE()), 23)
  GROUP BY f.COD_FORMULARIO, f.COD_LOJA
)
SELECT u.COD_FORMULARIO, u.COD_LOJA,
       a51.DES_ALTERNATIVA implementada,
       a100.DES_ALTERNATIVA motivo
FROM ultimaPorLoja u
JOIN bi_d_loja l ON l.COD_LOJA = u.COD_LOJA AND l.STATUS_LOJA = 'ATIVO' AND l.DES_REDE NOT LIKE '%TESTE%'
JOIN bi_f_formulario f51 ON f51.COD_FORMULARIO = u.COD_FORMULARIO AND f51.COD_LOJA = u.COD_LOJA AND f51.DATA = u.ultima_data AND f51.COD_PERGUNTA = 51
LEFT JOIN bi_d_formulario_alternativas a51 ON a51.bi_id_formulario_alternativas = f51.bi_id_formulario_alternativas
LEFT JOIN bi_f_formulario f100 ON f100.COD_FORMULARIO = u.COD_FORMULARIO AND f100.COD_LOJA = u.COD_LOJA AND f100.DATA = u.ultima_data AND f100.COD_PERGUNTA = 100
LEFT JOIN bi_d_formulario_alternativas a100 ON a100.bi_id_formulario_alternativas = f100.bi_id_formulario_alternativas
"@
$campanhaNomes = Invoke-Sql @"
SELECT COD_FORMULARIO, DES_FORMULARIO, CONVERT(varchar, DT_INICIAL, 23) dt_inicial
FROM bi_d_formulario
WHERE COD_FORMULARIO IN (
  SELECT COD_FORMULARIO FROM bi_f_formulario WHERE COD_PERGUNTA = 51
  GROUP BY COD_FORMULARIO HAVING MAX(DATA) >= CONVERT(varchar, DATEADD(day,-14,GETDATE()), 23)
)
"@
$campanhaNomesMap = @{}
foreach ($r in $campanhaNomes.Rows) { $campanhaNomesMap[(N $r.COD_FORMULARIO)] = @{ nome = (S $r.DES_FORMULARIO); inicio = (S $r.dt_inicial) } }
# "MATERIAL DESCARTADO PELA LOJA": o promotor implementou e a loja jogou o material fora
# depois - a campanha aconteceu, entao conta como implementada (marcada a parte no detalhe).
function Descartado([string]$motivo) { return $motivo -match 'DESCARTAD' }

$porCampanha = @{}
foreach ($r in $campanhaRespostas.Rows) {
    $cf = N $r.COD_FORMULARIO
    if (-not $porCampanha.ContainsKey($cf)) { $porCampanha[$cf] = @{ total = 0; sim = 0; descartadas = 0; motivos = @{} } }
    $porCampanha[$cf].total++
    $mo = S $r.motivo
    if ((S $r.implementada) -eq 'SIM') {
        $porCampanha[$cf].sim++
    } elseif (Descartado $mo) {
        $porCampanha[$cf].sim++
        $porCampanha[$cf].descartadas++
    } else {
        if ($mo) {
            if (-not $porCampanha[$cf].motivos.ContainsKey($mo)) { $porCampanha[$cf].motivos[$mo] = 0 }
            $porCampanha[$cf].motivos[$mo]++
        }
    }
}
$campanhas = New-Object System.Collections.ArrayList
foreach ($cf in $porCampanha.Keys) {
    $info = $porCampanha[$cf]
    $meta = $campanhaNomesMap[$cf]
    $piorMotivo = $null; $piorMotivoQtd = 0
    foreach ($mo in $info.motivos.Keys) { if ($info.motivos[$mo] -gt $piorMotivoQtd) { $piorMotivoQtd = $info.motivos[$mo]; $piorMotivo = $mo } }
    [void]$campanhas.Add([ordered]@{
        formulario     = $cf
        nome           = if ($meta) { $meta.nome } else { "Formulário $cf" }
        inicio         = if ($meta) { $meta.inicio } else { '' }
        lojasAvaliadas = $info.total
        implementadas  = $info.sim
        descartadas    = $info.descartadas
        taxa           = if ($info.total -gt 0) { [math]::Round($info.sim / $info.total, 4) } else { 0 }
        motivoMaisComum = $piorMotivo
    })
}
$campanhas = @($campanhas | Sort-Object -Property { $_.inicio } -Descending)

# ---------- 7b) campanhas: detalhe por loja + quantidades de MPDV implementado ----------
# cada campanha tem seu proprio conjunto de materiais (perguntas "QUANTIDADE IMPLEMENTADA
# DE X"), entao descobrimos as perguntas de cada formulario dinamicamente em vez de fixar
# uma lista - times de trade trocam os materiais a cada campanha nova.
Write-Host "  consultando detalhe de campanhas (lojas + MPDV)..."
$lojaInfoRows = Invoke-Sql @"
SELECT COD_LOJA, DES_REDE rede, NOM_FANTASIA nome, DES_UF uf
FROM bi_d_loja
WHERE STATUS_LOJA = 'ATIVO' AND DES_REDE NOT LIKE '%TESTE%'
"@
$lojaInfoMap = @{}
foreach ($r in $lojaInfoRows.Rows) { $lojaInfoMap[(N $r.COD_LOJA)] = @{ rede = (NomeRede (S $r.rede)); nome = [System.Net.WebUtility]::HtmlDecode((S $r.nome)); uf = (S $r.uf) } }

$respostaPorLoja = @{}
foreach ($r in $campanhaRespostas.Rows) {
    $cf = N $r.COD_FORMULARIO; $cl = N $r.COD_LOJA
    if (-not $respostaPorLoja.ContainsKey($cf)) { $respostaPorLoja[$cf] = @{} }
    $mo = S $r.motivo
    $desc = (S $r.implementada) -ne 'SIM' -and (Descartado $mo)
    $respostaPorLoja[$cf][$cl] = @{ implementada = $(if ($desc) { 'SIM' } else { S $r.implementada }); motivo = $mo; descartado = $desc }
}

$campanhasDetalhe = New-Object System.Collections.ArrayList
foreach ($cf in $porCampanha.Keys) {
    $meta = $campanhaNomesMap[$cf]
    $nomeCampanha = if ($meta) { $meta.nome } else { "Formulário $cf" }

    # bi_d_formulario_perguntas nao tem COD_FORMULARIO - descobre as perguntas de
    # quantidade desse formulario via bi_f_formulario (quem de fato foi respondido nele).
    $perguntasQtd = Invoke-Sql @"
SELECT DISTINCT f.COD_PERGUNTA, p.DES_PERGUNTA
FROM bi_f_formulario f
JOIN bi_d_formulario_perguntas p ON p.COD_PERGUNTA = f.COD_PERGUNTA
WHERE f.COD_FORMULARIO = $cf
  AND f.DATA >= CONVERT(varchar, DATEADD(day,-14,GETDATE()), 23)
  AND p.DES_PERGUNTA LIKE 'QUANTIDADE IMPLEMENTADA%'
"@
    $materiaisDef = New-Object System.Collections.ArrayList
    $codsPergunta = New-Object System.Collections.ArrayList
    foreach ($r in $perguntasQtd.Rows) {
        $cp = N $r.COD_PERGUNTA
        $label = (S $r.DES_PERGUNTA) -replace '(?i)^QUANTIDADE IMPLEMENTADA(\s+DE)?\s*', '' -replace '\?\s*$', ''
        $label = $label.Trim()
        [void]$materiaisDef.Add(@{ cod = $cp; label = $label })
        [void]$codsPergunta.Add("$cp")
    }
    if ($codsPergunta.Count -eq 0) { continue }
    $codsStr = ($codsPergunta -join ',')

    $qtdRows = Invoke-Sql @"
;WITH ultimaPorLoja AS (
  SELECT COD_LOJA, MAX(DATA) ultima_data
  FROM bi_f_formulario
  WHERE COD_FORMULARIO = $cf AND COD_PERGUNTA = 51
    AND DATA >= CONVERT(varchar, DATEADD(day,-14,GETDATE()), 23)
  GROUP BY COD_LOJA
)
SELECT f.COD_LOJA, f.COD_PERGUNTA, f.QTD
FROM bi_f_formulario f
JOIN ultimaPorLoja u ON u.COD_LOJA = f.COD_LOJA AND u.ultima_data = f.DATA
WHERE f.COD_FORMULARIO = $cf AND f.COD_PERGUNTA IN ($codsStr)
"@
    $qtdPorLoja = @{}
    foreach ($r in $qtdRows.Rows) {
        $loja = N $r.COD_LOJA
        if (-not $qtdPorLoja.ContainsKey($loja)) { $qtdPorLoja[$loja] = @{} }
        $qtdPorLoja[$loja][(N $r.COD_PERGUNTA)] = (N $r.QTD)
    }

    # fotos de execucao (pergunta tipo "photo", ex: "TIRE NO MIN 2 FOTOS AMPLAS DA
    # EXECUCAO") - a URL fica em bi_d_formulario_respostas.DES_RESPOSTA. Pega so as
    # mais recentes pra nao inchar o doc "meta" do Firestore com centenas de URLs.
    $perguntasFoto = Invoke-Sql @"
SELECT DISTINCT f.COD_PERGUNTA
FROM bi_f_formulario f
JOIN bi_d_formulario_perguntas p ON p.COD_PERGUNTA = f.COD_PERGUNTA
WHERE f.COD_FORMULARIO = $cf
  AND f.DATA >= CONVERT(varchar, DATEADD(day,-14,GETDATE()), 23)
  AND p.DES_TIPO_PERGUNTA = 'photo'
"@
    $codsFoto = @($perguntasFoto.Rows | ForEach-Object { N $_.COD_PERGUNTA })
    $fotosCampanha = New-Object System.Collections.ArrayList
    if ($codsFoto.Count -gt 0) {
        $codsFotoStr = ($codsFoto -join ',')
        $fotoRows = Invoke-Sql @"
SELECT TOP 30 f.COD_LOJA, f.DATA, r.DES_RESPOSTA url
FROM bi_f_formulario f
JOIN bi_d_formulario_respostas r ON r.bi_id_formulario_respostas = f.bi_id_formulario_respostas
WHERE f.COD_FORMULARIO = $cf AND f.COD_PERGUNTA IN ($codsFotoStr)
  AND f.DATA >= CONVERT(varchar, DATEADD(day,-14,GETDATE()), 23)
  AND r.DES_RESPOSTA IS NOT NULL AND r.DES_RESPOSTA <> ''
ORDER BY f.DATA DESC
"@
        foreach ($r in $fotoRows.Rows) {
            $loja = N $r.COD_LOJA
            $info = $lojaInfoMap[$loja]
            if (-not $info) { continue }
            [void]$fotosCampanha.Add([ordered]@{
                url  = (S $r.url)
                rede = $info.rede
                nome = $info.nome
                data = (S $r.DATA)
            })
        }
    }

    $lojasCampanha = New-Object System.Collections.ArrayList
    $respMap = $respostaPorLoja[$cf]
    foreach ($cl in $respMap.Keys) {
        $info = $lojaInfoMap[$cl]
        if (-not $info) { continue }
        $resp = $respMap[$cl]
        $materiais = [ordered]@{}
        foreach ($m in $materiaisDef) {
            $q = 0
            if ($qtdPorLoja.ContainsKey($cl) -and $qtdPorLoja[$cl].ContainsKey($m.cod)) { $q = $qtdPorLoja[$cl][$m.cod] }
            $materiais[$m.label] = $q
        }
        [void]$lojasCampanha.Add([ordered]@{
            rede         = $info.rede
            nome         = $info.nome
            uf           = $info.uf
            implementada = $resp.implementada
            descartado   = $resp.descartado
            motivo       = $resp.motivo
            materiais    = $materiais
        })
    }

    [void]$campanhasDetalhe.Add([ordered]@{
        formulario = $cf
        nome       = $nomeCampanha
        materiais  = @($materiaisDef | ForEach-Object { $_.label })
        lojas      = @($lojasCampanha)
        fotos      = @($fotosCampanha)
    })
}
$campanhasDetalhe = @($campanhasDetalhe | Sort-Object -Property {
    $m = $campanhaNomesMap[$_.formulario]; if ($m) { $m.inicio } else { '' }
} -Descending)

# ---------- 7c) ocorrencias em loja (formulario 4 "OCORRENCIAS") ----------
# so as descricoes (pergunta 11) dos ultimos 90 dias, sempre ao vivo (nao muda com o
# filtro de mes). Rede/loja/UF o front pega da carteira pelo cod, pra seguir os filtros.
Write-Host "  consultando ocorrencias..."
$ocorrRows = Invoke-Sql @"
SELECT f.COD_LOJA, CONVERT(varchar, f.DATA, 23) DATA, r.DES_RESPOSTA descricao, pes.NOM_PESSOA_COMPLETO promotor
FROM bi_f_formulario f
JOIN bi_d_formulario_respostas r ON r.bi_id_formulario_respostas = f.bi_id_formulario_respostas
LEFT JOIN bi_d_pessoa pes ON pes.COD_PESSOA = f.COD_PESSOA
WHERE f.COD_FORMULARIO = 4 AND f.COD_PERGUNTA = 11
  AND f.DATA >= CONVERT(varchar, DATEADD(day,-90,GETDATE()), 23)
  AND LTRIM(RTRIM(ISNULL(r.DES_RESPOSTA,''))) <> ''
ORDER BY f.DATA DESC
"@
$ocorrencias = New-Object System.Collections.ArrayList
foreach ($r in $ocorrRows.Rows) {
    $cl = N $r.COD_LOJA
    if (-not $lojaInfoMap.ContainsKey($cl)) { continue }
    [void]$ocorrencias.Add([ordered]@{
        cod       = "$cl"
        data      = (S $r.DATA)
        descricao = [System.Net.WebUtility]::HtmlDecode((S $r.descricao))
        promotor  = (S $r.promotor)
    })
}

# ---------- 7d) base de lojas com Nintendo (planilha do roteiro) ----------
# universo da cobertura = lojas atendidas + lojas da base nao visitadas. Pega a planilha
# mais recente da pasta do roteiro e cruza com o cadastro (tools\cruzar_base_lojas.py).
# Sem planilha/python, a cobertura volta a usar todas as lojas ativas do banco.
$pastaBaseLojas = 'K:\JOB 222100 - NINTENDO\2. Roteiro'
$arqCoordBase = Join-Path $PSScriptRoot '..\dados\base_lojas_coordenadas.xlsx'
$baseLojas = $null
$baseCod = @{}
try {
    # dados\base_lojas recebe versoes corrigidas mandadas direto (ex.: enderecos arrumados); vale a mais recente
    $pastasBase = @($pastaBaseLojas, (Join-Path $PSScriptRoot '..\dados\base_lojas')) | Where-Object { Test-Path -LiteralPath $_ }
    $arqBase = $pastasBase | ForEach-Object { Get-ChildItem -LiteralPath $_ -Recurse -Filter 'BASE GERAL DE LOJAS COM NINTENDO*.xlsx' -ErrorAction Stop } |
        Where-Object { $_.Name -notlike '~$*' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $arqBase) { throw "nenhuma planilha 'BASE GERAL DE LOJAS COM NINTENDO' em $pastaBaseLojas" }
    Write-Host "  cruzando base de lojas: $($arqBase.Name)..."
    $lojasCad = Invoke-Sql @"
SELECT COD_LOJA cod, NUM_CNPJ cnpj, DES_REDE rede, DES_BANDEIRA bandeira, NOM_FANTASIA nome, DES_CIDADE cidade, DES_UF uf
FROM bi_d_loja WHERE STATUS_LOJA = 'ATIVO' AND DES_REDE NOT LIKE '%TESTE%'
"@
    $tmpLojas = Join-Path $env:TEMP 'nintendo_lojas_cadastro.json'
    # fica em dados\ (local, NDA): o importador de estoque do varejo usa o mesmo de-para de lojas
    $tmpSaida = Join-Path $PSScriptRoot '..\dados\base_lojas_cruzamento.json'
    @($lojasCad.Rows | ForEach-Object { [ordered]@{
        cod = "$($_.cod)"; cnpj = (S $_.cnpj); rede = (S $_.rede); bandeira = (S $_.bandeira)
        nome = [System.Net.WebUtility]::HtmlDecode((S $_.nome)); cidade = (S $_.cidade); uf = (S $_.uf)
    } }) | ConvertTo-Json -Depth 3 | Set-Content -Encoding UTF8 $tmpLojas
    $argsPy = @((Join-Path $PSScriptRoot 'cruzar_base_lojas.py'), $arqBase.FullName, $tmpLojas, $tmpSaida)
    if (Test-Path $arqCoordBase) { $argsPy += (Resolve-Path $arqCoordBase).Path }
    & python @argsPy | ForEach-Object { Write-Host "    $_" }
    if ($LASTEXITCODE -ne 0) { throw "cruzar_base_lojas.py saiu com codigo $LASTEXITCODE" }
    $cruz = Get-Content -Raw -Encoding UTF8 $tmpSaida | ConvertFrom-Json
    # loja marcada como ONLINE na base (vende so online) nao entra na carteira fisica;
    # se for visitada, continua contando como atendida pelo proprio cadastro.
    foreach ($c in $cruz.casadas) { if (-not $c.online) { $baseCod["$($c.cod)"] = $true } }
    $semCadastro = @($cruz.semCadastro | Where-Object { -not $_.online })
    $baseLojas = [ordered]@{ arquivo = $arqBase.Name; linhas = $cruz.linhasBase; semCadastro = $semCadastro }
    Write-Host "    carteira da base sem lojas online: $($baseCod.Count) cadastradas + $($semCadastro.Count) fora do cadastro"
} catch {
    Write-Warning "base de lojas nao aplicada ($($_.Exception.Message)) - cobertura usa todas as lojas ativas do banco"
}

# ---------- monta 1 snapshot completo de loja, ancorado numa data "ref" ----------
function Get-SnapshotDoMes([datetime]$inicio, [datetime]$ref, [bool]$ehAtual) {
    $refStr = Dt $ref
    $inicioStr = Dt $inicio
    $mixIniStr = Dt ($ref.AddDays(-$janelaMixDias))
    $ultimaIniStr = Dt ($ref.AddDays(-$janelaVisitaDias))
    $asOf = $ref.AddDays(-1)

    Exec "IF OBJECT_ID('tempdb..#mix') IS NOT NULL DROP TABLE #mix"
    Exec "IF OBJECT_ID('tempdb..#ultima') IS NOT NULL DROP TABLE #ultima"
    Exec "IF OBJECT_ID('tempdb..#ultimo_sku') IS NOT NULL DROP TABLE #ultimo_sku"
    Exec "IF OBJECT_ID('tempdb..#semanas') IS NOT NULL DROP TABLE #semanas"

    # ---- roteiro do mes: soma das semanas de bi_f_efetividade dentro de [inicio, ref) ----
    # "atual" cobre do 1o dia do mes corrente ate hoje (mes em andamento); meses passados
    # cobrem o mes fechado inteiro. planejadoSemana/realizadoSemana viram CONTAGEM de
    # semanas (nao mais flag 0/1 de "semana atual").
    Exec @"
SELECT e.cod_loja, t.PrimeiroDiaSemana,
       MAX(e.planejado) AS planejado_semana,
       MAX(e.realizado) AS realizado_semana
INTO #semanas
FROM bi_f_efetividade e
JOIN bi_d_tempo t ON t.bi_id_tempo = e.bi_id_tempo
WHERE t.PrimeiroDiaSemana >= '$inicioStr' AND t.PrimeiroDiaSemana < '$refStr'
GROUP BY e.cod_loja, t.PrimeiroDiaSemana
"@
    $cobertura = Invoke-Sql @"
SELECT cod_loja,
       SUM(CASE WHEN planejado_semana > 0 THEN 1 ELSE 0 END) semanas_planejadas,
       SUM(CASE WHEN planejado_semana > 0 AND realizado_semana > 0 THEN 1 ELSE 0 END) semanas_realizadas
FROM #semanas
GROUP BY cod_loja
"@
    $coberturaMap = @{}
    foreach ($r in $cobertura.Rows) {
        $coberturaMap[(S $r.cod_loja)] = @{ planejado = (N $r.semanas_planejadas); realizado = (N $r.semanas_realizadas) }
    }

    Exec @"
SELECT bi_id_loja, bi_id_sku, MAX(DATA) AS ultima_presenca
INTO #mix
FROM bi_f_produto
WHERE FL_PRESENTE = 1 AND bi_id_sku > 3 AND DATA >= '$mixIniStr' AND DATA < '$refStr'
  -- dashboard e da Nintendo: produto PlayStation/Xbox fora do mix (nao conta como ruptura)
  AND bi_id_sku NOT IN (SELECT bi_id_sku FROM bi_d_sku WHERE DES_MARCA IN ('PLAYSTATION','XBOX'))
GROUP BY bi_id_loja, bi_id_sku
"@
    Exec @"
SELECT bi_id_loja, MAX(DATA) AS ultima_data
INTO #ultima
FROM bi_f_produto
WHERE DATA >= '$ultimaIniStr' AND DATA < '$refStr'
GROUP BY bi_id_loja
"@
    # colapsa possiveis registros duplicados de loja/sku no mesmo dia (mesma
    # visita gerando mais de uma linha) para 1 linha por loja/sku, senao o
    # SUM(FL_RUPTURA) fica maior que o COUNT DISTINCT de SKUs e a taxa passa de 100%
    Exec @"
SELECT p.bi_id_loja, p.bi_id_sku,
       MAX(p.FL_PRESENTE) AS presente,
       MAX(p.FL_RUPTURA) AS ruptura,
       MAX(p.FRENTES) AS frentes
INTO #ultimo_sku
FROM bi_f_produto p
JOIN #ultima u ON u.bi_id_loja = p.bi_id_loja AND u.ultima_data = p.DATA
WHERE p.bi_id_sku > 3
  AND p.bi_id_sku NOT IN (SELECT bi_id_sku FROM bi_d_sku WHERE DES_MARCA IN ('PLAYSTATION','XBOX'))
GROUP BY p.bi_id_loja, p.bi_id_sku
"@

    # ---- snapshot por loja (ruptura no mix, frentes, presenca) ----
    $snapshot = Invoke-Sql @"
SELECT l.COD_LOJA, l.DES_REDE, l.DES_BANDEIRA, l.NOM_FANTASIA, l.DES_CANAL,
       l.DES_CIDADE, l.DES_UF, l.DES_REGIAO, l.STATUS_LOJA,
       l.LATITUDE, l.LONGITUDE,
       u.ultima_data,
       pes.NOM_PESSOA_COMPLETO promotor, pes.SUPERVISOR supervisor,
       COALESCE(mixstats.itens_no_mix, 0) itens_no_mix,
       COALESCE(mixstats.rupturas_no_mix, 0) rupturas_no_mix,
       COALESCE(allstats.skus_presentes, 0) skus_presentes,
       COALESCE(allstats.frentes_total, 0) frentes_total,
       COALESCE(allstats.skus_avaliados_total, 0) skus_avaliados_total
FROM bi_d_loja l
LEFT JOIN #ultima u ON u.bi_id_loja = l.bi_id_loja
LEFT JOIN (
    SELECT p3.bi_id_loja, MIN(p3.COD_PESSOA) COD_PESSOA
    FROM bi_f_produto p3
    JOIN #ultima u3 ON u3.bi_id_loja = p3.bi_id_loja AND u3.ultima_data = p3.DATA
    GROUP BY p3.bi_id_loja
) pessoaUlt ON pessoaUlt.bi_id_loja = l.bi_id_loja
LEFT JOIN bi_d_pessoa pes ON pes.COD_PESSOA = pessoaUlt.COD_PESSOA
LEFT JOIN (
    SELECT us.bi_id_loja, COUNT(*) itens_no_mix, SUM(us.ruptura) rupturas_no_mix
    FROM #ultimo_sku us
    JOIN #mix mm ON mm.bi_id_loja = us.bi_id_loja AND mm.bi_id_sku = us.bi_id_sku
    GROUP BY us.bi_id_loja
) mixstats ON mixstats.bi_id_loja = l.bi_id_loja
LEFT JOIN (
    SELECT us.bi_id_loja,
           SUM(us.frentes) frentes_total,
           SUM(us.presente) skus_presentes,
           COUNT(*) skus_avaliados_total
    FROM #ultimo_sku us
    GROUP BY us.bi_id_loja
) allstats ON allstats.bi_id_loja = l.bi_id_loja
WHERE l.STATUS_LOJA = 'ATIVO' AND l.DES_REDE NOT LIKE '%TESTE%'
"@

    # ---- detalhe do mix sku a sku (em falta e presentes) na ultima visita ----
    # ultima_presenca = ultima vez que o sku foi visto na loja na janela do mix: pra sku em
    # falta, mostra ha quanto tempo esta faltando.
    $rupturaDetalhe = Invoke-Sql @"
SELECT l.COD_LOJA, s.DES_SKU, s.DES_CATEGORIA, s.DES_MARCA, us.frentes FRENTES, us.ruptura RUPTURA, us.presente PRESENTE,
       CONVERT(varchar, mm.ultima_presenca, 23) ULTIMA_PRESENCA
FROM #ultimo_sku us
JOIN #mix mm ON mm.bi_id_loja = us.bi_id_loja AND mm.bi_id_sku = us.bi_id_sku
JOIN bi_d_loja l ON l.bi_id_loja = us.bi_id_loja AND l.STATUS_LOJA = 'ATIVO' AND l.DES_REDE NOT LIKE '%TESTE%'
JOIN bi_d_sku s ON s.bi_id_sku = us.bi_id_sku
WHERE us.ruptura = 1 OR us.presente = 1
"@

    # ---- mobiliario (formulario 3: situacao, tipo, motivo ausencia, consultor), como estava ate "ref" ----
    # limitado a janela de $janelaVisitaDias dias (mesma janela de "ultima visita" usada pra
    # ruptura/presenca) - sem isso, uma resposta de mobiliario de meses atras contava como dado
    # "atual" pra sempre, distorcendo a nota de execucao de lojas ha muito sem visita.
    $display = Invoke-Sql @"
;WITH ult AS (
  SELECT COD_LOJA, MAX(DATA) AS ultima_data
  FROM bi_f_formulario WHERE COD_FORMULARIO = 3 AND COD_PERGUNTA = 9 AND DATA >= '$ultimaIniStr' AND DATA < '$refStr'
  GROUP BY COD_LOJA
),
base AS (
  SELECT f.COD_LOJA, f.DATA, f.COD_PESSOA, f.COD_PERGUNTA, a.DES_ALTERNATIVA
  FROM ult u
  JOIN bi_f_formulario f ON f.COD_LOJA = u.COD_LOJA AND f.DATA = u.ultima_data AND f.COD_FORMULARIO = 3 AND f.COD_PERGUNTA IN (9,13,228)
  LEFT JOIN bi_d_formulario_alternativas a ON a.bi_id_formulario_alternativas = f.bi_id_formulario_alternativas
)
SELECT b.COD_LOJA, MAX(b.DATA) DATA,
       MAX(CASE WHEN b.COD_PERGUNTA=9 THEN b.DES_ALTERNATIVA END) situacao,
       MAX(CASE WHEN b.COD_PERGUNTA=13 THEN b.DES_ALTERNATIVA END) tipo,
       MAX(CASE WHEN b.COD_PERGUNTA=228 THEN b.DES_ALTERNATIVA END) motivo,
       pes.NOM_PESSOA_COMPLETO consultor
FROM base b
LEFT JOIN bi_d_pessoa pes ON pes.COD_PESSOA = b.COD_PESSOA
GROUP BY b.COD_LOJA, pes.NOM_PESSOA_COMPLETO
"@

    $displayMap = @{}
    foreach ($r in $display.Rows) {
        $displayMap[(S $r.COD_LOJA)] = @{
            situacao = (S $r.situacao); data = (S $r.DATA); tipo = (S $r.tipo)
            motivo = (S $r.motivo); consultor = (S $r.consultor)
        }
    }

    # ---- treinamento em loja (formulario 1: "TREINAMENTO OTJ"), como estava ate "ref" ----
    # limitado a janela de $janelaVisitaDias dias (mesma janela de "ultima visita" usada pra
    # ruptura/mobiliario) - sem isso, um treinamento de meses/anos atras contava como dado
    # "atual" pra sempre, distorcendo a aba de Treinamentos com redes/lojas ha muito sem visita.
    $treinamento = Invoke-Sql @"
;WITH ult AS (
  SELECT COD_LOJA, MAX(DATA) AS ultima_data
  FROM bi_f_formulario WHERE COD_FORMULARIO = 1 AND COD_PERGUNTA = 1 AND DATA >= '$ultimaIniStr' AND DATA < '$refStr'
  GROUP BY COD_LOJA
),
base AS (
  SELECT f.COD_LOJA, f.DATA, f.COD_PERGUNTA, a.DES_ALTERNATIVA, f.QTD, r.DES_RESPOSTA
  FROM ult u
  JOIN bi_f_formulario f ON f.COD_LOJA = u.COD_LOJA AND f.DATA = u.ultima_data AND f.COD_FORMULARIO = 1 AND f.COD_PERGUNTA IN (1,2,3,203)
  LEFT JOIN bi_d_formulario_alternativas a ON a.bi_id_formulario_alternativas = f.bi_id_formulario_alternativas
  LEFT JOIN bi_d_formulario_respostas r ON r.bi_id_formulario_respostas = f.bi_id_formulario_respostas
)
SELECT b.COD_LOJA, MAX(b.DATA) DATA,
       MAX(CASE WHEN b.COD_PERGUNTA=1 THEN b.DES_ALTERNATIVA END) realizado,
       -- pergunta tipo "text": o numero digitado fica em DES_RESPOSTA (QTD e sempre 1)
       MAX(CASE WHEN b.COD_PERGUNTA=2 THEN TRY_CONVERT(int, LTRIM(RTRIM(b.DES_RESPOSTA))) END) qtdPessoas,
       MAX(CASE WHEN b.COD_PERGUNTA=3 THEN b.DES_ALTERNATIVA END) motivo,
       MAX(CASE WHEN b.COD_PERGUNTA=203 THEN b.DES_ALTERNATIVA END) nivelConhecimento
FROM base b
GROUP BY b.COD_LOJA
"@
    $treinamentoMap = @{}
    foreach ($r in $treinamento.Rows) {
        $treinamentoMap[(S $r.COD_LOJA)] = @{
            realizado = (S $r.realizado); data = (S $r.DATA); qtdPessoas = (N $r.qtdPessoas)
            motivo = (S $r.motivo); nivelConhecimento = (S $r.nivelConhecimento)
        }
    }

    # ---- demonstracao de produto (formulario 8), como estava ate "ref" ----
    # hora inicio/fim (perguntas 30/31) sao tipo "text" e o valor mora em
    # bi_d_formulario_respostas.DES_RESPOSTA (ex: "16:40"), nao em DES_ALTERNATIVA
    # limitado a janela de $janelaVisitaDias dias pelo mesmo motivo do treinamento acima.
    $demonstracao = Invoke-Sql @"
;WITH ult AS (
  SELECT COD_LOJA, MAX(DATA) AS ultima_data
  FROM bi_f_formulario WHERE COD_FORMULARIO = 8 AND COD_PERGUNTA = 28 AND DATA >= '$ultimaIniStr' AND DATA < '$refStr'
  GROUP BY COD_LOJA
),
base AS (
  SELECT f.COD_LOJA, f.DATA, f.COD_PERGUNTA, a.DES_ALTERNATIVA, f.QTD, r.DES_RESPOSTA
  FROM ult u
  JOIN bi_f_formulario f ON f.COD_LOJA = u.COD_LOJA AND f.DATA = u.ultima_data AND f.COD_FORMULARIO = 8 AND f.COD_PERGUNTA IN (28,29,30,31)
  LEFT JOIN bi_d_formulario_alternativas a ON a.bi_id_formulario_alternativas = f.bi_id_formulario_alternativas
  LEFT JOIN bi_d_formulario_respostas r ON r.bi_id_formulario_respostas = f.bi_id_formulario_respostas
)
SELECT b.COD_LOJA, MAX(b.DATA) DATA,
       MAX(CASE WHEN b.COD_PERGUNTA=28 THEN b.DES_ALTERNATIVA END) realizado,
       MAX(CASE WHEN b.COD_PERGUNTA=29 THEN TRY_CONVERT(int, LTRIM(RTRIM(b.DES_RESPOSTA))) END) qtdPessoas,
       MAX(CASE WHEN b.COD_PERGUNTA=30 THEN b.DES_RESPOSTA END) horaInicio,
       MAX(CASE WHEN b.COD_PERGUNTA=31 THEN b.DES_RESPOSTA END) horaFim
FROM base b
GROUP BY b.COD_LOJA
"@
    $demonstracaoMap = @{}
    foreach ($r in $demonstracao.Rows) {
        $hi = S $r.horaInicio; $hf = S $r.horaFim
        $duracaoMin = $null
        if ($hi -and $hf) {
            try {
                $tsIni = [datetime]::ParseExact($hi, 'HH:mm', $null)
                $tsFim = [datetime]::ParseExact($hf, 'HH:mm', $null)
                $diff = ($tsFim - $tsIni).TotalMinutes
                if ($diff -ge 0 -and $diff -le 600) { $duracaoMin = [math]::Round($diff, 0) }
            } catch {}
        }
        $demonstracaoMap[(S $r.COD_LOJA)] = @{
            realizado = (S $r.realizado); data = (S $r.DATA); qtdPessoas = (N $r.qtdPessoas); duracaoMin = $duracaoMin
        }
    }

    # ---- fotos da ultima visita: mobiliario (form 3, pergunta 10) e demonstracao (form 8, pergunta 213) ----
    $fotosDisplay = FotosUltimaVisita 3 9 10 $ultimaIniStr $refStr
    $fotosDemo = FotosUltimaVisita 8 28 213 $ultimaIniStr $refStr

    # ---- estoque de console NINTENDO no mix (pilar "Estoque console" da Loja Perfeita) ----
    # categoria CONSOLE inclui Playstation/Xbox - so NS1/NS2 contam aqui.
    $console = Invoke-Sql @"
SELECT l.COD_LOJA, COUNT(*) console_no_mix, SUM(us.ruptura) console_rupturas,
       SUM(CASE WHEN s.DES_MARCA = 'NS1' THEN 1 ELSE 0 END) console_ns1,
       SUM(CASE WHEN s.DES_MARCA = 'NS2' THEN 1 ELSE 0 END) console_ns2
FROM #ultimo_sku us
JOIN #mix mm ON mm.bi_id_loja = us.bi_id_loja AND mm.bi_id_sku = us.bi_id_sku
JOIN bi_d_loja l ON l.bi_id_loja = us.bi_id_loja
JOIN bi_d_sku s ON s.bi_id_sku = us.bi_id_sku
WHERE s.DES_CATEGORIA = 'CONSOLE' AND s.DES_MARCA IN ('NS1','NS2')
GROUP BY l.COD_LOJA
"@
    $consoleMap = @{}
    foreach ($r in $console.Rows) { $consoleMap[(S $r.COD_LOJA)] = @{ mix = (N $r.console_no_mix); rupturas = (N $r.console_rupturas); ns1 = (N $r.console_ns1); ns2 = (N $r.console_ns2) } }

    # ---- ativacao de MPDV: ultima resposta da pergunta 51 ("LOJA TEM IMPLEMENTACAO DE
    # CAMPANHA?", comum a todos os formularios de campanha) na janela de visita. Motivo
    # (pergunta 100) "MATERIAL DESCARTADO PELA LOJA" conta como implementada, igual nas campanhas ----
    $mpdv = Invoke-Sql @"
;WITH ult AS (
  SELECT COD_LOJA, MAX(DATA) ultima_data
  FROM bi_f_formulario WHERE COD_PERGUNTA = 51 AND DATA >= '$ultimaIniStr' AND DATA < '$refStr'
  GROUP BY COD_LOJA
)
SELECT u.COD_LOJA, u.ultima_data DATA,
       MAX(CASE WHEN a.DES_ALTERNATIVA = 'SIM' OR a100.DES_ALTERNATIVA LIKE '%DESCARTAD%' THEN 1 ELSE 0 END) implementada
FROM ult u
JOIN bi_f_formulario f ON f.COD_LOJA = u.COD_LOJA AND f.DATA = u.ultima_data AND f.COD_PERGUNTA = 51
LEFT JOIN bi_d_formulario_alternativas a ON a.bi_id_formulario_alternativas = f.bi_id_formulario_alternativas
LEFT JOIN bi_f_formulario f100 ON f100.COD_FORMULARIO = f.COD_FORMULARIO AND f100.COD_LOJA = u.COD_LOJA AND f100.DATA = u.ultima_data AND f100.COD_PERGUNTA = 100
LEFT JOIN bi_d_formulario_alternativas a100 ON a100.bi_id_formulario_alternativas = f100.bi_id_formulario_alternativas
GROUP BY u.COD_LOJA, u.ultima_data
"@
    $mpdvMap = @{}
    foreach ($r in $mpdv.Rows) { $mpdvMap[(S $r.COD_LOJA)] = @{ implementada = $(if ((N $r.implementada) -eq 1) { 'SIM' } else { "N$([char]0x00C3)O" }); data = (S $r.DATA) } }

    # ---- loja "atendida" = esta no roteiro (planejada) ou teve visita na janela ----
    # a base tem todas as lojas ATIVAS da Nintendo, inclusive as que o time de campo nao
    # cobre (ex: boa parte do cluster Carrefour) - essas nao entram nas abas de execucao.
    $roteiroJanela = Invoke-Sql @"
SELECT DISTINCT e.cod_loja
FROM bi_f_efetividade e
JOIN bi_d_tempo t ON t.bi_id_tempo = e.bi_id_tempo
WHERE e.planejado > 0 AND t.PrimeiroDiaSemana >= '$ultimaIniStr' AND t.PrimeiroDiaSemana < '$refStr'
"@
    $noRoteiro = @{}
    foreach ($r in $roteiroJanela.Rows) { $noRoteiro[(S $r.cod_loja)] = $true }

    # ultima semana com visita realizada no roteiro - entra na "ultima visita geral" junto
    # com as datas dos formularios. Sem isso, loja visitada so pra campanha/treinamento
    # aparecia como "nunca visitada" (ultimaVisita e so a do levantamento de estoque).
    $semanaRealizada = Invoke-Sql @"
SELECT e.cod_loja, CONVERT(varchar, MAX(t.PrimeiroDiaSemana), 23) ultima_semana
FROM bi_f_efetividade e
JOIN bi_d_tempo t ON t.bi_id_tempo = e.bi_id_tempo
WHERE e.realizado > 0 AND t.PrimeiroDiaSemana >= '$ultimaIniStr' AND t.PrimeiroDiaSemana < '$refStr'
GROUP BY e.cod_loja
"@
    $semanaRealizadaMap = @{}
    foreach ($r in $semanaRealizada.Rows) { $semanaRealizadaMap[(S $r.cod_loja)] = (S $r.ultima_semana) }

    # rupturaDetalhe = so os skus em falta (filtros e exportacao usam assim); mixPresente = o
    # resto do mix, pra ficha da loja na aba de ruptura
    $rupturaPorLoja = @{}
    $presentePorLoja = @{}
    foreach ($r in $rupturaDetalhe.Rows) {
        $k = S $r.COD_LOJA
        $item = @{ sku = (S $r.DES_SKU); categoria = (S $r.DES_CATEGORIA); marca = (S $r.DES_MARCA); frentes = (N $r.FRENTES) }
        $destino = $presentePorLoja
        if ((N $r.RUPTURA) -eq 1) { $item.ultimaPresenca = (S $r.ULTIMA_PRESENCA); $destino = $rupturaPorLoja }
        if (-not $destino.ContainsKey($k)) { $destino[$k] = New-Object System.Collections.ArrayList }
        [void]$destino[$k].Add($item)
    }

    $lojas = New-Object System.Collections.ArrayList
    foreach ($r in $snapshot.Rows) {
        $cod = S $r.COD_LOJA
        $itensNoMix = N $r.itens_no_mix
        $rupturas = N $r.rupturas_no_mix
        $taxaRuptura = 0
        if ($itensNoMix -gt 0) { $taxaRuptura = [math]::Round(($rupturas / $itensNoMix), 4) }

        $ultimaData = S $r.ultima_data
        $diasSemVisita = $null
        if ($ultimaData) {
            $diasSemVisita = [math]::Floor(($asOf - [datetime]$ultimaData).TotalDays)
        }

        $disp = $displayMap[$cod]
        $cob = $coberturaMap[$cod]
        $trein = $treinamentoMap[$cod]
        $demo = $demonstracaoMap[$cod]
        $cons = $consoleMap[$cod]
        $mp = $mpdvMap[$cod]
        $treinamentoDias = $null
        if ($trein -and $trein.data) { $treinamentoDias = [math]::Floor(($asOf - [datetime]$trein.data).TotalDays) }
        $atendida = [bool]($noRoteiro.ContainsKey($cod) -or $null -ne $diasSemVisita -or $disp -or $trein -or $demo -or $mp)
        $datasVisita = @($ultimaData, $semanaRealizadaMap[$cod]) + @($disp, $trein, $demo, $mp | Where-Object { $_ } | ForEach-Object { $_.data })
        $ultimaVisitaGeral = @($datasVisita | Where-Object { $_ } | ForEach-Object { "$_".Substring(0,10) } | Sort-Object -Descending)[0]
        $diasSemVisitaGeral = $null
        if ($ultimaVisitaGeral) { $diasSemVisitaGeral = [math]::Max(0, [math]::Floor(($asOf - [datetime]$ultimaVisitaGeral).TotalDays)) }

        [void]$lojas.Add([ordered]@{
            cod           = $cod
            rede          = (NomeRede (S $r.DES_REDE))
            bandeira      = (S $r.DES_BANDEIRA)
            # cadastro guarda alguns nomes com entidade HTML ("B&amp;B GAMES")
            nome          = [System.Net.WebUtility]::HtmlDecode((S $r.NOM_FANTASIA))
            canal         = (S $r.DES_CANAL)
            cidade        = (S $r.DES_CIDADE)
            uf            = (S $r.DES_UF)
            regiao        = (S $r.DES_REGIAO)
            promotor      = (S $r.promotor)
            supervisor    = (S $r.supervisor)
            lat           = (S $r.LATITUDE)
            lng           = (S $r.LONGITUDE)
            ultimaVisita  = $ultimaData
            diasSemVisita = $diasSemVisita
            itensNoMix    = $itensNoMix
            rupturas      = $rupturas
            taxaRuptura   = $taxaRuptura
            skusPresentes = (N $r.skus_presentes)
            frentesTotal  = (N $r.frentes_total)
            skusAvaliados = (N $r.skus_avaliados_total)
            display       = if ($disp) { $disp.situacao } else { $null }
            displayData   = if ($disp) { $disp.data } else { $null }
            displayTipo   = if ($disp) { $disp.tipo } else { $null }
            displayMotivo = if ($disp) { $disp.motivo } else { $null }
            displayConsultor = if ($disp) { $disp.consultor } else { $null }
            displayFotos  = if ($fotosDisplay.ContainsKey($cod)) { @($fotosDisplay[$cod]) } else { @() }
            planejadoSemana = if ($cob) { $cob.planejado } else { 0 }
            realizadoSemana = if ($cob) { $cob.realizado } else { 0 }
            treinamentoRealizado = if ($trein) { $trein.realizado } else { $null }
            treinamentoData      = if ($trein) { $trein.data } else { $null }
            treinamentoQtdPessoas = if ($trein) { $trein.qtdPessoas } else { 0 }
            treinamentoMotivo    = if ($trein) { $trein.motivo } else { $null }
            treinamentoNivel     = if ($trein) { $trein.nivelConhecimento } else { $null }
            demonstracaoRealizada = if ($demo) { $demo.realizado } else { $null }
            demonstracaoData      = if ($demo) { $demo.data } else { $null }
            demonstracaoQtdPessoas = if ($demo) { $demo.qtdPessoas } else { 0 }
            demonstracaoDuracaoMin = if ($demo) { $demo.duracaoMin } else { $null }
            demonstracaoFotos = if ($fotosDemo.ContainsKey($cod)) { @($fotosDemo[$cod]) } else { @() }
            treinamentoDias = $treinamentoDias
            consoleNoMix   = if ($cons) { $cons.mix } else { 0 }
            consoleRupturas = if ($cons) { $cons.rupturas } else { 0 }
            consoleNS1     = if ($cons) { $cons.ns1 } else { 0 }
            consoleNS2     = if ($cons) { $cons.ns2 } else { 0 }
            ultimaVisitaGeral  = $ultimaVisitaGeral
            diasSemVisitaGeral = $diasSemVisitaGeral
            mpdvImplementada = if ($mp) { $mp.implementada } else { $null }
            mpdvData       = if ($mp) { $mp.data } else { $null }
            atendida       = $atendida
            naBase         = $baseCod.ContainsKey($cod)
            rupturaDetalhe = if ($rupturaPorLoja.ContainsKey($cod)) { $rupturaPorLoja[$cod] } else { @() }
            mixPresente    = if ($presentePorLoja.ContainsKey($cod)) { $presentePorLoja[$cod] } else { @() }
        })
    }

    $resultado = @{ lojas = $lojas }

    if ($ehAtual) {
        # ---- tendencia semanal (8 semanas) de ruptura no mix - so faz sentido pro mes atual, e usa o #mix de agora ----
        Write-Host "  consultando tendencia semanal..."
        $tendencia = Invoke-Sql @"
;WITH dia AS (
    SELECT p.bi_id_loja, p.bi_id_sku, p.DATA, p.bi_id_tempo, p.COD_LOJA,
           MAX(p.FL_RUPTURA) AS ruptura
    FROM bi_f_produto p
    JOIN #mix m ON m.bi_id_loja = p.bi_id_loja AND m.bi_id_sku = p.bi_id_sku
    WHERE p.DATA >= CONVERT(varchar, DATEADD(day, -56, GETDATE()), 23)
    GROUP BY p.bi_id_loja, p.bi_id_sku, p.DATA, p.bi_id_tempo, p.COD_LOJA
)
SELECT t.ano_semana, t.PrimeiroDiaSemana,
       COUNT(DISTINCT dia.COD_LOJA) lojas_visitadas,
       SUM(dia.ruptura) rupturas,
       COUNT(*) itens
FROM dia
JOIN bi_d_tempo t ON t.bi_id_tempo = dia.bi_id_tempo
GROUP BY t.ano_semana, t.PrimeiroDiaSemana
ORDER BY t.PrimeiroDiaSemana
"@
        $tendenciaArr = New-Object System.Collections.ArrayList
        foreach ($r in $tendencia.Rows) {
            $itens = N $r.itens
            $taxa = 0
            if ($itens -gt 0) { $taxa = [math]::Round(((N $r.rupturas) / $itens), 4) }
            [void]$tendenciaArr.Add([ordered]@{
                semana = (S $r.ano_semana)
                inicio = (S $r.PrimeiroDiaSemana)
                lojasVisitadas = (N $r.lojas_visitadas)
                taxaRuptura = $taxa
            })
        }
        $resultado.tendencia = $tendenciaArr
    }

    return $resultado
}

# ---------- roda 1 snapshot por mes configurado (atual primeiro) ----------
$porMes = [ordered]@{}
$mesesResumo = New-Object System.Collections.ArrayList
$tendenciaFinal = $null
foreach ($mc in $mesesConfig) {
    Write-Host "  montando snapshot: $($mc.rotulo)..."
    $snap = Get-SnapshotDoMes -inicio $mc.inicio -ref $mc.ref -ehAtual $mc.ehAtual
    $porMes[$mc.chave] = @{ lojas = $snap.lojas }
    if ($mc.ehAtual) { $tendenciaFinal = $snap.tendencia }
    [void]$mesesResumo.Add([ordered]@{ chave = $mc.chave; rotulo = $mc.rotulo })
}

$cn.Close()

# ---------- monta estrutura em memoria ----------
Write-Host "  montando JSON..."

$dados = [ordered]@{
    geradoEm      = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    janelaMixDias = $janelaMixDias
    lojas         = $porMes['atual'].lojas
    tendencia     = $tendenciaFinal
    skus          = $skus
    meses         = $mesesResumo
    mesAtual      = 'atual'
    porMes        = $porMes
    campanhas     = $campanhas
    campanhasDetalhe = $campanhasDetalhe
    ocorrencias = @($ocorrencias)
    baseLojas = $baseLojas
}

$json = $dados | ConvertTo-Json -Depth 12 -Compress
$jsBlock = "const DADOS=$json;"

# ---------- injeta o bloco no HTML ----------
$html = [IO.File]::ReadAllText($htmlFile)
$startMarker = '/*__DADOS_START__*/'
$endMarker = '/*__DADOS_END__*/'
$wrapped = "$startMarker`n$jsBlock`n$endMarker"

if ($html.Contains($startMarker)) {
    $pattern = [regex]::Escape($startMarker) + '.*?' + [regex]::Escape($endMarker)
    $html = [regex]::Replace($html, $pattern, [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $wrapped }, [System.Text.RegularExpressions.RegexOptions]::Singleline)
} else {
    throw "Não encontrei os marcadores $startMarker / $endMarker no dashboard.html."
}

[IO.File]::WriteAllText($htmlFile, $html, (New-Object System.Text.UTF8Encoding($false)))

Write-Host "OK -> $htmlFile" -ForegroundColor Green
Write-Host ("  lojas ativas: {0} | meses no filtro: {1}" -f $porMes['atual'].lojas.Count, (($mesesResumo | ForEach-Object { $_.rotulo }) -join ', '))
