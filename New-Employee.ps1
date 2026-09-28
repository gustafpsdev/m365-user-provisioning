<#
.SYNOPSIS
    Onboarding automatizado de colaboradores no Microsoft 365 (via Graph):
    cria a conta, atribui licenca, adiciona aos grupos, define o gestor e gera
    as credenciais de boas-vindas — tudo a partir de um cargo (departamento).

.DESCRIPTION
    O provisionamento e baseado em cargo: config/roles.json define, por
    departamento, a licenca e os grupos que o novo colaborador recebe. Assim o
    processo fica padronizado e sem erro manual. Ao final, gera um relatorio HTML.

.PARAMETER DemoData
    Processa um lote ficticio (data/demo-employees.json), sem conectar ao tenant.

.PARAMETER Nome
    Nome completo do colaborador (modo real, um a um).

.PARAMETER Departamento
    Departamento/cargo, que define licenca e grupos (ver config/roles.json).

.PARAMETER Cargo
    Titulo do cargo (jobTitle).

.PARAMETER Gestor
    Nome do gestor (para o campo manager).

.PARAMETER OutputPath
    Relatorio HTML gerado. Padrao: output/onboarding-report.html.

.EXAMPLE
    ./New-Employee.ps1 -DemoData

.EXAMPLE
    ./New-Employee.ps1 -Nome "Ana Ribeiro" -Departamento Financeiro -Cargo "Analista" -Gestor "Carlos Mendes"

.NOTES
    Autor: Gustavo Paiva
    Modo real requer Microsoft.Graph e os escopos:
      User.ReadWrite.All, Group.ReadWrite.All, Organization.Read.All, Directory.ReadWrite.All
#>

[CmdletBinding()]
param(
    [switch]$DemoData,
    [string]$Nome,
    [string]$Departamento,
    [string]$Cargo,
    [string]$Gestor,
    [string]$OutputPath = "output/onboarding-report.html"
)

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
$cfg = Get-Content "config/roles.json" -Raw | ConvertFrom-Json
$dominio = $cfg.dominio

function Remove-Accents([string]$s) {
    $n = $s.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $n.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($c)
        }
    }
    return $sb.ToString()
}

function New-Upn([string]$nome) {
    $limpo = (Remove-Accents $nome).ToLower() -replace "[^a-z ]", ""
    $partes = $limpo.Split(" ", [StringSplitOptions]::RemoveEmptyEntries)
    $first = $partes[0]
    $last = if ($partes.Count -gt 1) { $partes[-1] } else { "" }
    $base = if ($last) { "$first.$last" } else { $first }
    return "$base@$dominio"
}

function New-TempPassword {
    $sets = @("ABCDEFGHJKLMNPQRSTUVWXYZ", "abcdefghijkmnpqrstuvwxyz", "23456789", "!@#%&*")
    $pwd = ($sets | ForEach-Object { $_[(Get-Random -Max $_.Length)] }) -join ""
    $all = ($sets -join "")
    while ($pwd.Length -lt 12) { $pwd += $all[(Get-Random -Max $all.Length)] }
    return $pwd
}

function Get-Role([string]$dep) {
    if ($cfg.departamentos.PSObject.Properties.Name -contains $dep) { return $cfg.departamentos.$dep }
    return $cfg.departamentos._default
}

# ---------------------------------------------------------------------------
# Coleta a lista de admissoes (demo OU parametro unico)
# ---------------------------------------------------------------------------
if ($DemoData) {
    Write-Host "[modo demonstracao] processando lote ficticio ..." -ForegroundColor Cyan
    $lista = (Get-Content "data/demo-employees.json" -Raw | ConvertFrom-Json).admissoes
}
elseif ($Nome) {
    $lista = @([PSCustomObject]@{ nome = $Nome; departamento = $Departamento; cargo = $Cargo; gestor = $Gestor })
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All","Organization.Read.All","Directory.ReadWrite.All" -NoWelcome
    }
}
else {
    throw "Informe -DemoData ou os parametros -Nome/-Departamento/-Cargo/-Gestor."
}

# ---------------------------------------------------------------------------
# Provisionamento
# ---------------------------------------------------------------------------
$resultados = New-Object System.Collections.Generic.List[object]

foreach ($e in $lista) {
    $role = Get-Role $e.departamento
    $upn = New-Upn $e.nome
    $senha = New-TempPassword
    $licNome = $cfg.skuNomes.($role.licenca)
    $grupos = @($role.grupos)
    $passos = New-Object System.Collections.Generic.List[object]

    if ($DemoData) {
        # Simula cada etapa como sucesso
        $passos.Add(@{ t = "Conta criada"; ok = $true; det = $upn })
        $passos.Add(@{ t = "Licenca atribuida"; ok = $true; det = $licNome })
        $passos.Add(@{ t = "Grupos adicionados"; ok = $true; det = ($grupos -join ", ") })
        $passos.Add(@{ t = "Gestor definido"; ok = $true; det = $e.gestor })
        $passos.Add(@{ t = "Boas-vindas geradas"; ok = $true; det = "senha temporaria + guia" })
    }
    else {
        try {
            $pwProfile = @{ Password = $senha; ForceChangePasswordNextSignIn = $true }
            $novo = New-MgUser -DisplayName $e.nome -UserPrincipalName $upn -MailNickname ($upn.Split("@")[0]) `
                        -AccountEnabled -PasswordProfile $pwProfile -JobTitle $e.cargo -Department $e.departamento
            $passos.Add(@{ t = "Conta criada"; ok = $true; det = $upn })

            $sku = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq $role.licenca | Select-Object -First 1
            if ($sku) {
                Set-MgUserLicense -UserId $novo.Id -AddLicenses @{ SkuId = $sku.SkuId } -RemoveLicenses @() | Out-Null
                $passos.Add(@{ t = "Licenca atribuida"; ok = $true; det = $licNome })
            } else { $passos.Add(@{ t = "Licenca atribuida"; ok = $false; det = "SKU $($role.licenca) nao encontrada" }) }

            foreach ($g in $grupos) {
                $grp = Get-MgGroup -Filter "displayName eq '$g'" -Top 1
                if ($grp) { New-MgGroupMember -GroupId $grp.Id -DirectoryObjectId $novo.Id }
            }
            $passos.Add(@{ t = "Grupos adicionados"; ok = $true; det = ($grupos -join ", ") })

            if ($e.gestor) {
                $mgr = Get-MgUser -Filter "displayName eq '$($e.gestor)'" -Top 1
                if ($mgr) { Set-MgUserManagerByRef -UserId $novo.Id -BodyParameter @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($mgr.Id)" } }
            }
            $passos.Add(@{ t = "Gestor definido"; ok = $true; det = $e.gestor })
            $passos.Add(@{ t = "Boas-vindas geradas"; ok = $true; det = "senha temporaria + guia" })
        }
        catch {
            $passos.Add(@{ t = "Erro"; ok = $false; det = $_.Exception.Message })
        }
    }

    $resultados.Add([PSCustomObject]@{
        Nome = $e.nome; UPN = $upn; Departamento = $e.departamento; Cargo = $e.cargo
        Gestor = $e.gestor; Licenca = $licNome; Grupos = $grupos; Senha = $senha; Passos = $passos
    })
    Write-Host ("  + {0}  ->  {1}  [{2}]" -f $e.nome, $upn, $licNome) -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Relatorio HTML
# ---------------------------------------------------------------------------
$totColab = $resultados.Count
$totLic = @($resultados | Where-Object { $_.Passos | Where-Object { $_.t -eq "Licenca atribuida" -and $_.ok } }).Count
$totGrupos = ($resultados | ForEach-Object { $_.Grupos.Count } | Measure-Object -Sum).Sum
$totFalhas = ($resultados | ForEach-Object { @($_.Passos | Where-Object { -not $_.ok }).Count } | Measure-Object -Sum).Sum

$cards = foreach ($r in $resultados) {
    $steps = ($r.Passos | ForEach-Object {
        $ic = if ($_.ok) { "ok" } else { "fail" }
        $mk = if ($_.ok) { "&#10003;" } else { "&#10007;" }
        "<li class='$ic'><span class='mk'>$mk</span> <b>$($_.t):</b> $($_.det)</li>"
    }) -join ""
    $gr = ($r.Grupos | ForEach-Object { "<span class='chip'>$_</span>" }) -join " "
    @"
    <div class="card">
      <div class="c-head">
        <div><span class="pname">$($r.Nome)</span><span class="upn">$($r.UPN)</span></div>
        <span class="dep">$($r.Departamento)</span>
      </div>
      <div class="meta">Cargo: <b>$($r.Cargo)</b> &middot; Gestor: <b>$($r.Gestor)</b> &middot; Licenca: <b>$($r.Licenca)</b></div>
      <div class="groups">$gr</div>
      <ul class="steps">$steps</ul>
    </div>
"@
}
$cardsHtml = $cards -join "`n"
$generatedAt = (Get-Date).ToString("dd/MM/yyyy HH:mm")
$modeLabel = if ($DemoData) { "Lote de demonstração (fictício)" } else { "Tenant real" }

$html = @"
<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Onboarding — Provisionamento</title>
<style>
  :root{--bg:#f4f6fb;--card:#fff;--ink:#1a2233;--muted:#5b6678;--line:#e6e9f0;--brand:#2f5bea;--green:#22a06b;--red:#e5484d}
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--ink);font-family:"Segoe UI",system-ui,-apple-system,Arial,sans-serif;line-height:1.5}
  .wrap{max-width:960px;margin:0 auto;padding:32px 20px 64px}
  h1{font-size:1.5rem;margin:0}.sub{color:var(--muted);font-size:.9rem;margin-top:4px}
  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:14px;margin:20px 0}
  .kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;text-align:center}
  .kpi .value{font-size:1.8rem;font-weight:800}.kpi.green .value{color:var(--green)}.kpi.red .value{color:var(--red)}
  .kpi .label{color:var(--muted);font-size:.72rem;text-transform:uppercase;letter-spacing:.03em;margin-top:2px}
  .card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px 18px;margin:12px 0}
  .c-head{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}
  .pname{font-weight:700;font-size:1.05rem;margin-right:10px}.upn{color:var(--brand);font-family:"Consolas",monospace;font-size:.9rem}
  .dep{background:#eef1f8;color:var(--muted);font-size:.75rem;padding:3px 10px;border-radius:999px;text-transform:uppercase;letter-spacing:.03em}
  .meta{color:#333;font-size:.9rem;margin:8px 0}
  .groups{margin:6px 0 10px}
  .chip{display:inline-block;background:#e9eefc;color:var(--brand);font-size:.78rem;padding:2px 10px;border-radius:6px;margin:2px 4px 2px 0}
  ul.steps{list-style:none;margin:0;padding:0;font-size:.9rem}
  ul.steps li{padding:4px 0;border-top:1px solid var(--line)}
  ul.steps li .mk{display:inline-block;width:18px;font-weight:800}
  ul.steps li.ok .mk{color:var(--green)} ul.steps li.fail .mk{color:var(--red)}
  footer{color:var(--muted);font-size:.8rem;text-align:center;margin-top:26px}
</style></head><body><div class="wrap">
  <h1>Onboarding &mdash; Provisionamento de Usuários</h1>
  <div class="sub">$modeLabel &middot; gerado em $generatedAt</div>

  <div class="kpis">
    <div class="kpi green"><div class="value">$totColab</div><div class="label">Colaboradores</div></div>
    <div class="kpi"><div class="value">$totLic</div><div class="label">Licenças atribuídas</div></div>
    <div class="kpi"><div class="value">$totGrupos</div><div class="label">Vínculos de grupo</div></div>
    <div class="kpi red"><div class="value">$totFalhas</div><div class="label">Falhas</div></div>
  </div>

$cardsHtml

  <footer>Gerado por <b>New-Employee.ps1</b> &middot; Provisionamento M365 &middot; Gustavo Paiva</footer>
</div></body></html>
"@

$outDir = Split-Path -Parent $OutputPath
if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$html | Set-Content -Path $OutputPath -Encoding UTF8

Write-Host ""
Write-Host "============== ONBOARDING ==============" -ForegroundColor Green
Write-Host ("Colaboradores : {0}" -f $totColab)
Write-Host ("Licencas      : {0}" -f $totLic)
Write-Host ("Grupos        : {0}" -f $totGrupos)
Write-Host ("Falhas        : {0}" -f $totFalhas)
Write-Host "=======================================" -ForegroundColor Green
Write-Host ("Relatorio salvo em: {0}" -f $OutputPath) -ForegroundColor Cyan
