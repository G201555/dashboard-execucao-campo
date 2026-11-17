# ============================================================
# Pipeline diario do Dashboard Nintendo: atualiza os dados de campo,
# recarrega o arquivo semanal de estoque (se houver um novo), publica
# no Firestore e faz o deploy do Hosting - tudo sem intervencao manual.
# Pensado pra rodar via Tarefa Agendada do Windows todo dia as 8h.
#
# Credenciais do Firebase ficam em tools\.cred\, criptografadas com
# DPAPI do usuario do Windows que criou o arquivo - so esse usuario,
# nessa maquina, consegue descriptografar. Nunca ficam em texto puro.
#
# As regras do Firestore so ficam com escrita liberada durante a
# republicacao (poucos segundos); um bloco finally garante que elas
# voltam a ficar bloqueadas mesmo se algo falhar no meio do caminho.
# ============================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$logDir = Join-Path $root 'logs'
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
$logFile = Join-Path $logDir ("atualizar_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd_HHmm'))

function Log([string]$msg) {
    $linha = "[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $msg
    Write-Host $linha
    Add-Content -Path $logFile -Value $linha
}

function Get-Credencial {
    $emailPath = Join-Path $PSScriptRoot '.cred\firebase_email.txt'
    $senhaPath = Join-Path $PSScriptRoot '.cred\firebase_senha.secure'
    if (-not (Test-Path $emailPath) -or -not (Test-Path $senhaPath)) {
        throw "Credencial nao encontrada em tools\.cred\ - rode a configuracao inicial antes de agendar."
    }
    $email = Get-Content $emailPath -Raw
    $secure = Get-Content $senhaPath | ConvertTo-SecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    $senha = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    return @{ email = $email; senha = $senha }
}

function Set-RegraEscrita([bool]$liberar) {
    $rulesPath = Join-Path $root 'firestore.rules'
    $conteudo = [IO.File]::ReadAllText($rulesPath)
    $novo = if ($liberar) { 'allow write: if autorizado();' } else { 'allow write: if false;' }
    $conteudo = [regex]::Replace($conteudo, 'allow write: if (autorizado\(\)|false);', $novo)
    [IO.File]::WriteAllText($rulesPath, $conteudo, (New-Object System.Text.UTF8Encoding($false)))
}

function Deploy([string]$alvo) {
    Push-Location $root
    # o firebase escreve progresso no stderr; com 'Stop' o PowerShell 5.1 trata isso como erro
    # e o pipeline abortava no meio do deploy do hosting ("ERRO: hosting: upload complete")
    $eapAnterior = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $saida = & firebase deploy --only $alvo 2>&1 | Out-String
        Log $saida
        if ($LASTEXITCODE -ne 0) { throw "firebase deploy --only $alvo falhou (exit $LASTEXITCODE)" }
    } finally {
        $ErrorActionPreference = $eapAnterior
        Pop-Location
    }
}

# a tarefa roda as 8h, muitas vezes antes da VPN conectar: espera o SQL Server responder
# (tenta a cada 10 min por ate 3h) em vez de falhar e deixar o dash sem atualizar o dia todo
function Aguardar-Vpn {
    $tentativas = 18
    for ($i = 1; $i -le $tentativas; $i++) {
        $ok = Test-NetConnection SEU_SERVIDOR_SQL -Port 1433 -InformationLevel Quiet -WarningAction SilentlyContinue
        if ($ok) { return }
        Log "SQL Server inacessivel (VPN desconectada?) - tentativa $i de $tentativas, nova tentativa em 10 min..."
        if ($i -lt $tentativas) { Start-Sleep -Seconds 600 }
    }
    throw "SQL Server SEU_SERVIDOR_SQL inacessivel por 3h - conecte a VPN e rode o pipeline de novo"
}

try {
    Log "===== iniciando pipeline diario ====="
    Aguardar-Vpn

    Log "1/5 atualizando dados de campo (atualizar_dashboard.ps1)..."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'atualizar_dashboard.ps1') 2>&1 | ForEach-Object { Log $_ }
    if ($LASTEXITCODE -ne 0) { throw "atualizar_dashboard.ps1 falhou (exit $LASTEXITCODE)" }

    Log "2/5 atualizando estoque/vendas semanais (importar_estoque_semanal.ps1)..."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'importar_estoque_semanal.ps1') 2>&1 | ForEach-Object { Log $_ }
    if ($LASTEXITCODE -ne 0) { throw "importar_estoque_semanal.ps1 falhou (exit $LASTEXITCODE)" }

    $cred = Get-Credencial

    Log "3/5 liberando escrita no Firestore e publicando dados..."
    Set-RegraEscrita $true
    Deploy 'firestore:rules'
    try {
        Push-Location $root
        $saidaNode = & node (Join-Path $PSScriptRoot 'publicar_firestore.js') $cred.email $cred.senha 2>&1 | Out-String
        Log $saidaNode
        if ($LASTEXITCODE -ne 0) { throw "publicar_firestore.js falhou (exit $LASTEXITCODE)" }
    } finally {
        Pop-Location
    }

    Log "4/5 bloqueando escrita no Firestore de novo..."
    Set-RegraEscrita $false
    Deploy 'firestore:rules,hosting'

    Log "5/5 pipeline concluido com sucesso."
} catch {
    Log "ERRO: $($_.Exception.Message)"
    # garante que a escrita nao fica liberada mesmo se algo quebrou no meio
    try { Set-RegraEscrita $false; Deploy 'firestore:rules' } catch { Log "ERRO ao re-bloquear regras: $($_.Exception.Message)" }
    exit 1
}
