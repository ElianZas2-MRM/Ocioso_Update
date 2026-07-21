# Diagnostico de conectividad a LambdaTest.
# Corre en cualquier Windows sin Python. Genera Diagnostico_LambdaTest_resultado.txt
# al lado de este archivo. Ese .txt es el que hay que mandar para analizar.

$ErrorActionPreference = "Continue"
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$out = Join-Path $dir "Diagnostico_LambdaTest_resultado.txt"
$hostname = "hub.lambdatest.com"
$L = New-Object System.Collections.ArrayList

function Log($m) { [void]$L.Add($m); Write-Host $m }

Log "=================================================="
Log " DIAGNOSTICO LAMBDATEST"
Log " Fecha : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Log " PC    : $env:COMPUTERNAME   Usuario: $env:USERNAME"
Log "=================================================="

# ---------- 1. Credenciales ----------
Log ""
Log "--- 1. CREDENCIALES ---"
$credFile = $null
foreach ($c in @((Join-Path $dir "lambdatest_credentials.txt"),
                 (Join-Path (Split-Path -Parent $dir) "lambdatest_credentials.txt"))) {
    if (Test-Path $c) { $credFile = $c; break }
}
$user = ""; $key = ""
if ($credFile) {
    Log "Archivo: $credFile"
    foreach ($line in (Get-Content $credFile)) {
        if ($line -match '^\s*username\s*=\s*(.+)$')   { $user = $Matches[1].Trim() }
        if ($line -match '^\s*access_key\s*=\s*(.+)$') { $key  = $Matches[1].Trim() }
    }
    if ($user -eq "" -or $key -eq "") { Log "PROBLEMA: falta username o access_key." }
    elseif ($user -eq "TU_USUARIO" -or $key -eq "TU_ACCESS_KEY") {
        Log "PROBLEMA: son las credenciales PLANTILLA, no las reales."
    } else {
        Log "username  : $user"
        Log "access_key: OK (largo $($key.Length), termina en ...$($key.Substring($key.Length-4)))"
    }
} else {
    Log "PROBLEMA: no se encontro lambdatest_credentials.txt"
}

# ---------- 2. DNS ----------
Log ""
Log "--- 2. DNS ---"
try {
    $ips = [System.Net.Dns]::GetHostAddresses($hostname) | ForEach-Object { $_.IPAddressToString }
    Log "$hostname resuelve a: $($ips -join ', ')"
} catch {
    Log "FALLA DNS: $($_.Exception.Message)"
}

# ---------- 3. TCP 443 ----------
Log ""
Log "--- 3. CONEXION TCP AL PUERTO 443 ---"
try {
    $tcp = New-Object System.Net.Sockets.TcpClient
    $ar = $tcp.BeginConnect($hostname, 443, $null, $null)
    if ($ar.AsyncWaitHandle.WaitOne(8000, $false) -and $tcp.Connected) {
        Log "TCP 443: CONECTA OK"
        $tcp.EndConnect($ar)
    } else {
        Log "TCP 443: NO CONECTA (timeout 8s) -> firewall o red bloqueando"
    }
    $tcp.Close()
} catch {
    Log "TCP 443: FALLA -> $($_.Exception.Message)"
}

# ---------- 4. Certificado TLS (detecta inspeccion corporativa) ----------
Log ""
Log "--- 4. CERTIFICADO TLS (clave para detectar proxy corporativo) ---"
try {
    $tcp2 = New-Object System.Net.Sockets.TcpClient($hostname, 443)
    $cb = { param($sn, $cert, $chain, $errors) return $true }
    $ssl = New-Object System.Net.Security.SslStream($tcp2.GetStream(), $false, $cb)
    $ssl.AuthenticateAsClient($hostname)
    $cert = $ssl.RemoteCertificate
    Log "Emisor  : $($cert.Issuer)"
    Log "Sujeto  : $($cert.Subject)"
    Log "Protocolo: $($ssl.SslProtocol)"
    if ($cert.Issuer -match "Amazon|DigiCert|Let's Encrypt|Sectigo|GlobalSign|Google Trust") {
        Log "=> Emisor publico normal. NO hay inspeccion TLS."
    } else {
        Log "=> ATENCION: emisor NO publico. La red esta INTERCEPTANDO el TLS"
        Log "   (proxy/antivirus corporativo). Esta es la causa mas probable del fallo:"
        Log "   Python valida contra certifi y rechaza este certificado."
    }
    $ssl.Close(); $tcp2.Close()
} catch {
    Log "No se pudo inspeccionar el certificado: $($_.Exception.Message)"
}

# ---------- 5. Proxy configurado ----------
Log ""
Log "--- 5. PROXY DEL SISTEMA ---"
try {
    $proxy = [System.Net.WebRequest]::GetSystemWebProxy().GetProxy("https://$hostname")
    if ($proxy.AbsoluteUri -like "*$hostname*") { Log "Proxy sistema: (ninguno, salida directa)" }
    else { Log "Proxy sistema: $($proxy.AbsoluteUri)  <-- hay proxy" }
} catch { Log "Proxy sistema: no se pudo determinar" }
foreach ($v in @("HTTP_PROXY","HTTPS_PROXY","NO_PROXY")) {
    $val = [Environment]::GetEnvironmentVariable($v)
    if ($val) { Log "Variable $v = $val" }
}

# ---------- 6. HTTPS sin credenciales ----------
Log ""
Log "--- 6. HTTPS AL HUB (sin credenciales) ---"
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $r = Invoke-WebRequest -Uri "https://$hostname/wd/hub/status" -TimeoutSec 20 -UseBasicParsing
    Log "HTTP $($r.StatusCode) -> el hub responde. Salida a internet OK."
} catch {
    Log "FALLA: $($_.Exception.Message)"
    if ($_.Exception.Message -match "SSL|trust|certificad|secure channel") {
        Log "=> Error de certificado: confirma inspeccion TLS corporativa."
    }
}

# ---------- 7. HTTPS con credenciales ----------
Log ""
Log "--- 7. AUTENTICACION CONTRA LAMBDATEST ---"
if ($user -and $key -and $user -ne "TU_USUARIO") {
    try {
        $pair = "$($user):$($key)"
        $b64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($pair))
        $r2 = Invoke-WebRequest -Uri "https://api.lambdatest.com/automation/api/v1/builds?limit=1" `
              -Headers @{ Authorization = "Basic $b64" } -TimeoutSec 20 -UseBasicParsing
        Log "HTTP $($r2.StatusCode) -> credenciales VALIDAS y API alcanzable."
    } catch {
        $code = ""
        if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
        Log "FALLA (HTTP $code): $($_.Exception.Message)"
        if ($code -eq 401) { Log "=> 401: usuario o access key INCORRECTOS." }
    }
} else {
    Log "Omitido: no hay credenciales reales cargadas."
}

Log ""
Log "=================================================="
Log " FIN. Mandar este archivo:"
Log " $out"
Log "=================================================="

$L | Out-File -FilePath $out -Encoding utf8
Write-Host ""
Write-Host "Resultado guardado en: $out" -ForegroundColor Green
Write-Host "Presiona Enter para cerrar..."
[void][Console]::ReadLine()
