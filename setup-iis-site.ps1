# Script para configurar sitio IIS para Excel Add-in
# Debe ejecutarse como Administrador

Import-Module WebAdministration

# Variables
$siteName = "ExcelAddinAzurriga"
$sitePort = 8443
$sitePath = "C:\excel-addin\ExcelRestAdding\dist"
$appPoolName = "ExcelAddinAppPool"

# Crear Application Pool
if (!(Test-Path "IIS:\AppPools\$appPoolName")) {
    Write-Host "Creando Application Pool: $appPoolName"
    New-WebAppPool -Name $appPoolName
    Set-ItemProperty "IIS:\AppPools\$appPoolName" -Name managedRuntimeVersion -Value ""
    Write-Host "Application Pool creado." -ForegroundColor Green
} else {
    Write-Host "Application Pool ya existe." -ForegroundColor Yellow
}

# Eliminar sitio si ya existe
if (Test-Path "IIS:\Sites\$siteName") {
    Write-Host "Eliminando sitio existente: $siteName"
    Remove-WebSite -Name $siteName
}

# Crear sitio web
Write-Host "Creando sitio web: $siteName"
New-WebSite -Name $siteName `
    -Port $sitePort `
    -PhysicalPath $sitePath `
    -ApplicationPool $appPoolName `
    -Force

Write-Host "Sitio web creado correctamente." -ForegroundColor Green
Write-Host "URL: http://localhost:$sitePort" -ForegroundColor Cyan

# Configurar tipos MIME si es necesario
Write-Host "Configurando tipos MIME..."
Add-WebConfigurationProperty -PSPath "IIS:\Sites\$siteName" `
    -Filter "//staticContent" `
    -Name "." `
    -Value @{fileExtension='.json';mimeType='application/json'} `
    -ErrorAction SilentlyContinue

Write-Host "`nPara agregar HTTPS, necesitarás un certificado SSL." -ForegroundColor Yellow
Write-Host "Puedes usar el certificado de desarrollo que ya tienes en el proyecto." -ForegroundColor Yellow
