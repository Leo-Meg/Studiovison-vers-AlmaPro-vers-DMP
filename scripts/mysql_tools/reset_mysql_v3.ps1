# reset_mysql_v3.ps1
# VERSION CORRIGEE - utilise mysqld-nt.exe (nom reel sur MySQL 5.0 Windows)
# Executer en ADMINISTRATEUR sur la VM serveur AlmaPro

$MYSQL_DIR  = "C:\Program Files\MySQL\MySQL Server 5.0\bin"
$MYSQLDNT   = "$MYSQL_DIR\mysqld-nt.exe"
$MYSQL_CLI  = "$MYSQL_DIR\mysql.exe"
$MYSQLADMIN = "$MYSQL_DIR\mysqladmin.exe"

Write-Host "=== Reset mot de passe MySQL (mysqld-nt) ===" -ForegroundColor Cyan
Write-Host ""

# Verification des fichiers
Write-Host "[0] Verification des fichiers..." -ForegroundColor Yellow
if (-not (Test-Path $MYSQLDNT))  { Write-Host "ERREUR: $MYSQLDNT introuvable" -ForegroundColor Red; exit 1 }
if (-not (Test-Path $MYSQL_CLI)) { Write-Host "ERREUR: $MYSQL_CLI introuvable" -ForegroundColor Red; exit 1 }
Write-Host "  OK - mysqld-nt.exe trouve" -ForegroundColor Green
Write-Host "  OK - mysql.exe trouve" -ForegroundColor Green

# Etape 1 : arreter le service MySQL
Write-Host "[1] Arret du service MySQL..." -ForegroundColor Yellow
$svc = Get-Service -Name "MySQL" -ErrorAction SilentlyContinue
if ($svc) {
    net stop MySQL 2>$null
    Start-Sleep -Seconds 3
}
# Tuer tout process mysqld-nt residuel
Stop-Process -Name "mysqld-nt" -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Write-Host "  OK" -ForegroundColor Green

# Etape 2 : demarrer mysqld-nt sans authentification
Write-Host "[2] Demarrage mysqld-nt --skip-grant-tables..." -ForegroundColor Yellow
$proc = Start-Process `
    -FilePath $MYSQLDNT `
    -ArgumentList "--skip-grant-tables", "--skip-networking" `
    -NoNewWindow -PassThru
Start-Sleep -Seconds 5
if ($proc.HasExited) {
    Write-Host "  ERREUR: mysqld-nt s'est arrete immediatement" -ForegroundColor Red
    Write-Host "  Essai sans --skip-networking..."
    $proc = Start-Process `
        -FilePath $MYSQLDNT `
        -ArgumentList "--skip-grant-tables" `
        -NoNewWindow -PassThru
    Start-Sleep -Seconds 5
}
Write-Host "  OK - PID $($proc.Id)" -ForegroundColor Green

# Etape 3 : supprimer le mot de passe root
Write-Host "[3] Suppression du mot de passe root..." -ForegroundColor Yellow
$sql = "UPDATE mysql.user SET Password='' WHERE User='root'; FLUSH PRIVILEGES;"
$r = & $MYSQL_CLI -u root -e $sql 2>&1
Write-Host "  Resultat: $r"
Write-Host "  OK" -ForegroundColor Green

# Etape 4 : arreter mysqld-nt temporaire
Write-Host "[4] Arret mysqld-nt temporaire..." -ForegroundColor Yellow
Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
Stop-Process -Name "mysqld-nt" -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3
Write-Host "  OK" -ForegroundColor Green

# Etape 5 : relancer le service MySQL normal
Write-Host "[5] Relance du service MySQL..." -ForegroundColor Yellow
net start MySQL
Start-Sleep -Seconds 5

# Etape 6 : verification connexion sans mot de passe
Write-Host "[6] Verification..." -ForegroundColor Yellow
$test = & $MYSQL_CLI -u root --connect-timeout=5 -e "SELECT 'CONNEXION_OK' AS r;" 2>&1
if ($test -match "CONNEXION_OK") {
    Write-Host ""
    Write-Host "  SUCCES - root sans mot de passe fonctionne !" -ForegroundColor Green
    Write-Host ""
    Write-Host "Depuis le poste client :" -ForegroundColor Cyan
    Write-Host "  python diag_mysql.py" -ForegroundColor Cyan
} else {
    Write-Host "  Resultat brut: $test" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Essai methode alternative via init-file..." -ForegroundColor Yellow

    # Methode alternative : init-file
    $initFile = "$env:TEMP\mysql_reset.sql"
    "UPDATE mysql.user SET Password='' WHERE User='root'; FLUSH PRIVILEGES;" | Out-File -FilePath $initFile -Encoding ASCII

    net stop MySQL 2>$null
    Stop-Process -Name "mysqld-nt" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    $proc2 = Start-Process `
        -FilePath $MYSQLDNT `
        -ArgumentList "--init-file=$initFile" `
        -NoNewWindow -PassThru
    Start-Sleep -Seconds 6
    Stop-Process -Id $proc2.Id -Force -ErrorAction SilentlyContinue
    Stop-Process -Name "mysqld-nt" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    net start MySQL
    Start-Sleep -Seconds 4

    $test2 = & $MYSQL_CLI -u root --connect-timeout=5 -e "SELECT 'OK' AS r;" 2>&1
    if ($test2 -match "OK") {
        Write-Host "  SUCCES via init-file !" -ForegroundColor Green
    } else {
        Write-Host "  Resultat: $test2" -ForegroundColor Red
        Write-Host ""
        Write-Host "  SOLUTION ALTERNATIVE : recuperer le mot de passe depuis AlmaPro" -ForegroundColor Yellow
        Write-Host "  Voir instructions dans README" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "=== Termine ===" -ForegroundColor Cyan
