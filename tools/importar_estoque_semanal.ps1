# ============================================================
# Importador do arquivo semanal de estoque loja a loja (cliente)
# Le um CSV local (nunca sobe pra lugar nenhum - fica so na sua
# maquina, conforme a NDA) e injeta estoque/venda/giro/cobertura
# no dashboard.html, cruzando com a ruptura de campo ja calculada
# pelo atualizar_dashboard.ps1 (aba "Execução × vendas").
#
# Uso:
#   1) Rode primeiro .\atualizar_dashboard.ps1 (precisa do DADOS
#      de ruptura/loja já injetado no HTML para o cruzamento).
#   2) Coloque o arquivo do cliente em:
#        dados\estoque_semanal.csv
#      com as colunas (nomes exatos, cabecalho na 1a linha):
#        cod_loja, cod_sku, data_referencia, estoque_unidades,
#        venda_semana_unidades, venda_semana_anterior_unidades
#   3) Rode:
#        .\importar_estoque_semanal.ps1
#
# Se dados\estoque_semanal.csv nao existir, usa o arquivo de
# EXEMPLO (dados\estoque_semanal.EXEMPLO.csv) so pra validar o
# pipeline - o dashboard mostra um aviso deixando claro que são
# dados fake, nunca deixa parecer dado real do cliente.
#
# Giro = venda da semana / estoque atual.
# Cobertura (dias) = estoque atual / venda media diaria (venda/7).
# Excesso de estoque = estoque no top 25% da rede E giro no
# bottom 25% da rede (ambos entre lojas com estoque > 0).
# ============================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$htmlFile = Join-Path $root 'dashboard.html'

# relatorios de estoque que as redes mandam (Carrefour, Sam's, Havan...) em dados\estoque_varejo:
# quando existem, sao eles que alimentam Vendas & Giro (importar_estoque_varejo.py)
$pastaVarejo = Join-Path $root 'dados\estoque_varejo'
if (Test-Path $pastaVarejo) {
    if (Get-ChildItem -Path $pastaVarejo -Filter '*.xlsx' | Where-Object { $_.Name -notlike '~$*' }) {
        & python (Join-Path $PSScriptRoot 'importar_estoque_varejo.py')
        exit $LASTEXITCODE
    }
}
$csvReal = Join-Path $root 'dados\estoque_semanal.csv'
$csvExemplo = Join-Path $root 'dados\estoque_semanal.EXEMPLO.csv'

$csvFile = if (Test-Path $csvReal) { $csvReal } else { $csvExemplo }
$isExemplo = ($csvFile -eq $csvExemplo)

if (-not (Test-Path $csvFile)) {
    throw "Nenhum arquivo encontrado. Coloque o arquivo semanal em dados\estoque_semanal.csv (colunas: cod_loja,cod_sku,data_referencia,estoque_unidades,venda_semana_unidades,venda_semana_anterior_unidades)."
}

Write-Host ("Importando {0}{1}" -f $csvFile, $(if ($isExemplo) { ' (ARQUIVO DE EXEMPLO - ainda sem dado real do cliente)' } else { '' })) -ForegroundColor Magenta

$rows = Import-Csv -Path $csvFile -Delimiter ','
if (-not $rows -or $rows.Count -eq 0) { throw "CSV vazio: $csvFile" }

$colunasObrigatorias = 'cod_loja', 'cod_sku', 'data_referencia', 'estoque_unidades', 'venda_semana_unidades', 'venda_semana_anterior_unidades'
$colunasArquivo = $rows[0].psobject.Properties.Name
foreach ($c in $colunasObrigatorias) {
    if ($colunasArquivo -notcontains $c) { throw "Coluna obrigatoria ausente no CSV: $c (colunas encontradas: $($colunasArquivo -join ', '))" }
}

function ToNum([string]$v) {
    if ([string]::IsNullOrWhiteSpace($v)) { return 0 }
    return [double]($v -replace ',', '.')
}

# ---------- carrega metadados de loja (rede/bandeira/uf/nome/ruptura) do dashboard.html ----------
$lojaMeta = @{}
if (Test-Path $htmlFile) {
    $htmlContent = [IO.File]::ReadAllText($htmlFile)
    $m = [regex]::Match($htmlContent, [regex]::Escape('/*__DADOS_START__*/') + '(.*?)' + [regex]::Escape('/*__DADOS_END__*/'), [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if ($m.Success) {
        $bloco = $m.Groups[1].Value
        $jm = [regex]::Match($bloco, 'const DADOS=(.*);')
        if ($jm.Success) {
            $dadosObj = $jm.Groups[1].Value | ConvertFrom-Json
            foreach ($lj in $dadosObj.lojas) {
                $lojaMeta[[string]$lj.cod] = @{
                    rede = $lj.rede; bandeira = $lj.bandeira; uf = $lj.uf; nome = $lj.nome; taxaRuptura = $lj.taxaRuptura
                }
            }
        }
    }
}
if ($lojaMeta.Count -eq 0) {
    Write-Host "  aviso: não achei DADOS de ruptura no dashboard.html ainda - rode atualizar_dashboard.ps1 primeiro para o cruzamento de execução x vendas funcionar." -ForegroundColor Yellow
}

# ---------- monta itens ----------
$itens = foreach ($r in $rows) {
    $estoque = ToNum $r.estoque_unidades
    $venda = ToNum $r.venda_semana_unidades
    $vendaAnt = ToNum $r.venda_semana_anterior_unidades
    [pscustomobject]@{
        cod_loja      = "$($r.cod_loja)".Trim()
        cod_sku       = "$($r.cod_sku)".Trim()
        estoque       = $estoque
        venda         = $venda
        vendaAnterior = $vendaAnt
        giro          = if ($estoque -gt 0) { [math]::Round($venda / $estoque, 3) } else { $null }
        coberturaDias = if ($venda -gt 0) { [math]::Round($estoque / ($venda / 7), 1) } else { $null }
    }
}

# ---------- agrega por loja ----------
$porLoja = $itens | Group-Object cod_loja | ForEach-Object {
    $g = $_.Group
    $estoqueTotal = ($g | Measure-Object estoque -Sum).Sum
    $vendaTotal = ($g | Measure-Object venda -Sum).Sum
    $vendaAntTotal = ($g | Measure-Object vendaAnterior -Sum).Sum
    [pscustomobject]@{
        cod_loja           = $_.Name
        estoqueTotal       = $estoqueTotal
        vendaTotal         = $vendaTotal
        vendaAnteriorTotal = $vendaAntTotal
        giro               = if ($estoqueTotal -gt 0) { [math]::Round($vendaTotal / $estoqueTotal, 3) } else { $null }
        coberturaDias      = if ($vendaTotal -gt 0) { [math]::Round($estoqueTotal / ($vendaTotal / 7), 1) } else { $null }
        wow                = if ($vendaAntTotal -gt 0) { [math]::Round((($vendaTotal - $vendaAntTotal) / $vendaAntTotal), 3) } else { $null }
        itens              = $g
    }
}

# ---------- quartis para flag de excesso de estoque ----------
function Percentil($arr, [double]$p) {
    if (-not $arr -or $arr.Count -eq 0) { return 0 }
    $idx = [math]::Floor(($arr.Count - 1) * $p)
    return $arr[$idx]
}
$estoques = @($porLoja | Where-Object { $_.estoqueTotal -gt 0 } | Select-Object -ExpandProperty estoqueTotal | Sort-Object)
$giros = @($porLoja | Where-Object { $null -ne $_.giro } | Select-Object -ExpandProperty giro | Sort-Object)
$estoqueP75 = Percentil $estoques 0.75
$giroP25 = Percentil $giros 0.25

# ---------- monta JSON ----------
$lojasVendas = foreach ($l in $porLoja) {
    $meta = $lojaMeta[$l.cod_loja]
    $excesso = ($l.estoqueTotal -ge $estoqueP75) -and ($null -ne $l.giro) -and ($l.giro -le $giroP25)
    [ordered]@{
        cod           = $l.cod_loja
        nome          = if ($meta) { $meta.nome } else { "Loja $($l.cod_loja)" }
        rede          = if ($meta) { $meta.rede } else { '' }
        bandeira      = if ($meta) { $meta.bandeira } else { '' }
        uf            = if ($meta) { $meta.uf } else { '' }
        taxaRuptura   = if ($meta) { $meta.taxaRuptura } else { $null }
        estoque       = $l.estoqueTotal
        venda         = $l.vendaTotal
        vendaAnterior = $l.vendaAnteriorTotal
        giro          = $l.giro
        coberturaDias = $l.coberturaDias
        wow           = $l.wow
        excesso       = $excesso
        itens         = $l.itens | ForEach-Object {
            [ordered]@{ sku = $_.cod_sku; estoque = $_.estoque; venda = $_.venda; giro = $_.giro; coberturaDias = $_.coberturaDias }
        }
    }
}

$dadosVendas = [ordered]@{
    geradoEm = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    exemplo  = $isExemplo
    arquivo  = (Split-Path -Leaf $csvFile)
    lojas    = $lojasVendas
}

$jsonV = $dadosVendas | ConvertTo-Json -Depth 10 -Compress
$jsBlockV = "const DADOS_VENDAS=$jsonV;"

# ---------- injeta no HTML ----------
$htmlContent = [IO.File]::ReadAllText($htmlFile)
$startV = '/*__VENDAS_START__*/'
$endV = '/*__VENDAS_END__*/'
$wrappedV = "$startV`n$jsBlockV`n$endV"

if ($htmlContent.Contains($startV)) {
    $pattern = [regex]::Escape($startV) + '.*?' + [regex]::Escape($endV)
    $htmlContent = [regex]::Replace($htmlContent, $pattern, [System.Text.RegularExpressions.MatchEvaluator]{ param($mm) $wrappedV }, [System.Text.RegularExpressions.RegexOptions]::Singleline)
} else {
    throw "Não encontrei os marcadores $startV / $endV no dashboard.html (precisa atualizar o HTML para a versão com a aba Vendas & Giro dinâmica)."
}

[IO.File]::WriteAllText($htmlFile, $htmlContent, (New-Object System.Text.UTF8Encoding($false)))

Write-Host "OK -> $htmlFile" -ForegroundColor Green
Write-Host ("  lojas com estoque/venda: {0} | excesso de estoque: {1}" -f $porLoja.Count, (@($lojasVendas | Where-Object { $_.excesso }).Count))
if ($isExemplo) { Write-Host "  ATENCAO: dados de EXEMPLO - troque por dados\estoque_semanal.csv com o arquivo real do cliente." -ForegroundColor Yellow }
