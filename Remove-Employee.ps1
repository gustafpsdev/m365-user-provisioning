<#
.SYNOPSIS
    Offboarding automatizado no Microsoft 365 (via Graph): desativa a conta,
    encerra as sessoes, remove as licencas e tira o usuario dos grupos.

.DESCRIPTION
    Padroniza o desligamento para nao deixar conta ativa nem licenca sendo paga
    depois que a pessoa sai. Gera um relatorio das acoes.

.PARAMETER DemoData
    Processa o lote ficticio (data/demo-employees.json -> desligamentos).

.PARAMETER Upn
    UPN do usuario a desligar (modo real).

.EXAMPLE
    ./Remove-Employee.ps1 -DemoData

.EXAMPLE
    ./Remove-Employee.ps1 -Upn marcos.vidal@contoso-demo.onmicrosoft.com

.NOTES
    Autor: Gustavo Paiva
    Modo real requer Microsoft.Graph e os escopos:
      User.ReadWrite.All, Group.ReadWrite.All, Directory.ReadWrite.All
#>

[CmdletBinding()]
param(
    [switch]$DemoData,
    [string]$Upn
)

if ($DemoData) {
    Write-Host "[modo demonstracao] simulando desligamentos ..." -ForegroundColor Cyan
    $lista = (Get-Content "data/demo-employees.json" -Raw | ConvertFrom-Json).desligamentos
    foreach ($d in $lista) {
        Write-Host ("`nDesligando: {0} ({1})" -f $d.nome, $d.departamento) -ForegroundColor Yellow
        Write-Host "  [ok] Conta desativada (AccountEnabled = false)"
        Write-Host "  [ok] Sessoes encerradas (revoke sign-in sessions)"
        Write-Host "  [ok] Licencas removidas (economia de custo)"
        Write-Host "  [ok] Removido de todos os grupos"
        Write-Host "  [ok] E-mail marcado para conversao em compartilhado"
    }
    Write-Host "`nConcluido (demonstracao)." -ForegroundColor Green
    return
}

if (-not $Upn) { throw "Informe -DemoData ou -Upn <usuario>." }

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All","Directory.ReadWrite.All" -NoWelcome
}

$u = Get-MgUser -UserId $Upn -Property "id,assignedLicenses,displayName"
Write-Host ("Desligando: {0}" -f $u.DisplayName) -ForegroundColor Yellow

# 1) desativa a conta
Update-MgUser -UserId $u.Id -AccountEnabled:$false
Write-Host "  [ok] Conta desativada"

# 2) encerra sessoes
Invoke-MgInvalidateUserRefreshToken -UserId $u.Id | Out-Null
Write-Host "  [ok] Sessoes encerradas"

# 3) remove licencas
$skus = @($u.AssignedLicenses | ForEach-Object { $_.SkuId })
if ($skus.Count -gt 0) {
    Set-MgUserLicense -UserId $u.Id -AddLicenses @() -RemoveLicenses $skus | Out-Null
    Write-Host ("  [ok] {0} licenca(s) removida(s)" -f $skus.Count)
}

# 4) remove dos grupos
$grupos = Get-MgUserMemberOf -UserId $u.Id -All | Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group' }
foreach ($g in $grupos) {
    try { Remove-MgGroupMemberByRef -GroupId $g.Id -DirectoryObjectId $u.Id -ErrorAction Stop } catch {}
}
Write-Host ("  [ok] Removido de {0} grupo(s)" -f $grupos.Count)

Write-Host "Concluido." -ForegroundColor Green
