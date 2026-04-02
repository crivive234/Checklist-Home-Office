#Requires -Version 5.0
# OTD Americas - Recopilacion de datos para Alistamiento Home Office

$ProgressPreference = "SilentlyContinue"

$data = [ordered]@{}
$data["generadoEn"] = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
$data["fecha"]      = (Get-Date).ToString("yyyy-MM-dd")

$data["hostname"]            = $env:COMPUTERNAME
$data["marca"]               = ""
$data["modelo"]              = ""
$data["serialNumber"]        = ""
$data["procesador"]          = ""
$data["nucleos"]             = ""
$data["ram"]                 = ""
$data["ramTipo"]             = ""
$data["ramFrecuencia"]       = ""
$data["discos"]              = @()
$data["biosVersion"]         = ""
$data["biosDate"]            = ""
$data["osVersion"]           = ""
$data["winRelease"]          = ""
$data["osArch"]              = ""
$data["drivers"]             = @()
$data["tarjetasRed"]         = @()
$data["pingResultado"]       = "fail"
$data["pingMs"]              = 0
$data["pingPerdida"]         = 4
$data["winUpdate"]           = "noRevisado"
$data["winUpdatePendientes"] = -1
$data["winLicencia"]         = "desconocido"
$data["winLicenciaTipo"]     = ""
$data["winLicenciaCanal"]    = ""
$data["programas"]           = @()

# ------------------------------------------------------------
# 1. HARDWARE
# ------------------------------------------------------------
Write-Host "1/5 Recopilando hardware y discos..." -ForegroundColor Cyan

$cs   = Get-CimInstance Win32_ComputerSystem  -ErrorAction SilentlyContinue
$os   = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
$cpu  = Get-CimInstance Win32_Processor       -ErrorAction SilentlyContinue | Select-Object -First 1
$bios = Get-CimInstance Win32_BIOS            -ErrorAction SilentlyContinue

if ($cs) {
    $data["marca"]  = $cs.Manufacturer.Trim()
    $data["modelo"] = $cs.Model.Trim()
}

if ($bios) {
    $data["serialNumber"] = $bios.SerialNumber.Trim()
    $data["biosVersion"]  = $bios.SMBIOSBIOSVersion.Trim()

    $biosDate = ""
    $biosRaw  = "$($bios.ReleaseDate)"

    # Método 1 (formato WMI)
    if ($biosRaw -match "^\d{14}") {
        $biosDate = ([System.Management.ManagementDateTimeConverter]::ToDateTime($bios.ReleaseDate)).ToString("yyyy-MM-dd")
    }

    # 🔧 Método 2 (REEMPLAZADO - sin TryParse)
    if (-not $biosDate -and $bios.ReleaseDate) {
        try {
            $parsed = [datetime]$biosRaw
            $biosDate = $parsed.ToString("yyyy-MM-dd")
        } catch {
            # no hacer nada
        }
    }

    # Método 3 (registro)
    if (-not $biosDate) {
        $regBios = Get-ItemProperty "HKLM:\HARDWARE\DESCRIPTION\System\BIOS" -ErrorAction SilentlyContinue
        if ($regBios -and $regBios.BIOSReleaseDate) {
            $biosDate = $regBios.BIOSReleaseDate
        }
    }

    $data["biosDate"] = $biosDate
}

if ($cpu) {
    $data["procesador"] = ($cpu.Name -replace "\s+", " ").Trim()
    $data["nucleos"]    = "$($cpu.NumberOfCores) nucleos / $($cpu.NumberOfLogicalProcessors) hilos"
}

$ramMem = Get-CimInstance Win32_PhysicalMemory -ErrorAction SilentlyContinue
if ($ramMem) {
    $ramBytes = ($ramMem | Measure-Object -Property Capacity -Sum).Sum
    $data["ram"] = "$([math]::Round($ramBytes / 1GB)) GB"

    # Tipo de RAM (SMBIOSMemoryType)
    $firstModule = $ramMem | Select-Object -First 1
    $tipoNum = $firstModule.SMBIOSMemoryType
    $tipoMap = @{
        20 = "DDR"; 21 = "DDR2"; 22 = "DDR2 FB-DIMM"; 24 = "DDR3";
        26 = "DDR4"; 34 = "DDR5"; 0 = "Desconocido"
    }
    $data["ramTipo"] = if ($tipoMap.ContainsKey([int]$tipoNum)) { $tipoMap[[int]$tipoNum] } else { "DDR ($tipoNum)" }

    # Frecuencia de RAM (Speed en MHz)
    $freqs = $ramMem | Where-Object { $_.Speed } | Select-Object -ExpandProperty Speed | Sort-Object -Unique
    $data["ramFrecuencia"] = if ($freqs) { ($freqs | ForEach-Object { "$_ MHz" }) -join " / " } else { "" }
}

if ($os) {
    $data["osVersion"] = "$($os.Caption) (Build $($os.BuildNumber))"
    $data["osArch"]    = $os.OSArchitecture
}

$regWin = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue
if ($regWin) {
    if ($regWin.DisplayVersion) {
        $data["winRelease"] = $regWin.DisplayVersion
    } elseif ($regWin.ReleaseId) {
        $data["winRelease"] = $regWin.ReleaseId
    }
}

# Discos
$discosRaw = Get-PhysicalDisk -ErrorAction SilentlyContinue
if ($discosRaw) {
    $discosList = $discosRaw | Sort-Object DeviceId | ForEach-Object {
        $d = $_

        # Tipo legible
        $tipoMap = @{ 3 = "HDD"; 4 = "SSD"; 5 = "SCM" }
        # MediaType puede ser numérico o string según el proveedor WMI
        $mediaRaw = "$($d.MediaType)"
        $tipoStr = switch -Regex ($mediaRaw) {
            "^3$"          { "HDD" }
            "^4$"          { "SSD" }
            "^5$"          { "SCM" }
            "HDD"          { "HDD" }
            "SSD"          { "SSD" }
            default        { "Desconocido" }
        }

        # Detectar M.2 / NVMe por nombre del bus o modelo
        $busType = ""
        try { $busType = $d.BusType } catch {}
        if ($busType -eq "NVMe" -or ($d.FriendlyName -match "NVMe|M\.2|PCIe")) {
            $tipoStr = "M.2 NVMe"
        } elseif ($busType -eq "SATA" -and $tipoStr -eq "SSD") {
            $tipoStr = "SSD SATA"
        } elseif ($busType -eq "SATA" -and $tipoStr -eq "HDD") {
            $tipoStr = "HDD SATA"
        }

        # Capacidad en GB / TB
        $bytes = [long]$d.Size
        $capStr = if ($bytes -ge 1TB) {
            "$([math]::Round($bytes / 1TB, 1)) TB"
        } else {
            "$([math]::Round($bytes / 1GB)) GB"
        }

        [ordered]@{
            modelo    = if ($d.FriendlyName) { $d.FriendlyName.Trim() } else { "Desconocido" }
            tipo      = $tipoStr
            capacidad = $capStr
            estado    = if ($d.HealthStatus) { $d.HealthStatus } else { "" }
        }
    }
    $data["discos"] = @($discosList)
    $cnt = $data["discos"].Count
    Write-Host "OK $cnt disco(s) detectado(s)" -ForegroundColor Green
} else {
    Write-Host "AVISO No se pudieron obtener discos fisicos" -ForegroundColor Yellow
}

Write-Host "OK Hardware recopilado - Version Windows: $($data[`"winRelease`"])" -ForegroundColor Green

# ------------------------------------------------------------
# 2. DRIVERS
# ------------------------------------------------------------
Write-Host "2/5 Recopilando drivers..." -ForegroundColor Cyan

$driversRaw = Get-CimInstance Win32_PnPSignedDriver -ErrorAction SilentlyContinue
if ($driversRaw) {
    $driversList = $driversRaw |
        Where-Object { $_.DriverVersion -and $_.DeviceName -and $_.DeviceClass } |
        Sort-Object DeviceClass, DeviceName |
        ForEach-Object {
            [ordered]@{
                nombre  = $_.DeviceName.Trim()
                clase   = $_.DeviceClass.Trim()
                version = $_.DriverVersion.Trim()
            }
        }
    $data["drivers"] = @($driversList)
    $cnt = $data["drivers"].Count
    Write-Host "OK $cnt drivers encontrados" -ForegroundColor Green
} else {
    Write-Host "ERROR No se pudieron obtener drivers" -ForegroundColor Red
}

# ------------------------------------------------------------
# 3. RED
# ------------------------------------------------------------
Write-Host "3/5 Recopilando tarjetas de red..." -ForegroundColor Cyan

$adaptadoresRaw = Get-NetAdapter -ErrorAction SilentlyContinue
if ($adaptadoresRaw) {
    $adaptadoresList = $adaptadoresRaw | Sort-Object Name | ForEach-Object {
        $adap = $_
        $ip = (Get-NetIPAddress -InterfaceIndex $adap.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1).IPAddress
        $velText = "N/A"
        if ($adap.LinkSpeed -and "$($adap.LinkSpeed)" -match "^\d+$") {
            $velText = "$([math]::Round([int64]$adap.LinkSpeed / 1MB)) Mbps"
        } elseif ($adap.LinkSpeed) {
            $velText = "$($adap.LinkSpeed)"
        }
        [ordered]@{
            nombre      = $adap.Name
            descripcion = $adap.InterfaceDescription
            mac         = $adap.MacAddress -replace "-", ":"
            estado      = $adap.Status
            velocidad   = $velText
            ip          = if ($ip) { $ip } else { "" }
        }
    }
    $data["tarjetasRed"] = @($adaptadoresList)
    $cnt = $data["tarjetasRed"].Count
    Write-Host "OK $cnt adaptadores encontrados" -ForegroundColor Green
} else {
    Write-Host "ERROR No se pudieron obtener adaptadores de red" -ForegroundColor Red
}

$pingResult = Test-Connection 8.8.8.8 -Count 4 -ErrorAction SilentlyContinue
if ($pingResult) {
    $data["pingResultado"] = "ok"
    $data["pingMs"]        = [math]::Round(($pingResult | Measure-Object ResponseTime -Average).Average, 1)
    $data["pingPerdida"]   = 4 - $pingResult.Count
    $ms = $data["pingMs"]
    Write-Host "OK Ping correcto ($ms ms)" -ForegroundColor Green
} else {
    Write-Host "ERROR Sin conectividad" -ForegroundColor Red
}

# ------------------------------------------------------------
# 4. WINDOWS UPDATE
# ------------------------------------------------------------
Write-Host "4/5 Verificando Windows Update..." -ForegroundColor Cyan

$wuSession = New-Object -ComObject Microsoft.Update.Session -ErrorAction SilentlyContinue
if ($wuSession) {
    $wuSearcher = $wuSession.CreateUpdateSearcher()
    $wuResult   = $wuSearcher.Search("IsInstalled=0 and IsHidden=0 and Type='Software'")
    if ($wuResult) {
        $count = $wuResult.Updates.Count
        if ($count -eq 0) {
            $data["winUpdate"] = "actualizado"
        } else {
            $data["winUpdate"] = "pendiente"
        }
        $data["winUpdatePendientes"] = $count
        $wu  = $data["winUpdate"]
        Write-Host "OK Windows Update: $wu ($count pendientes)" -ForegroundColor Green
    }
} else {
    Write-Host "AVISO Windows Update no disponible" -ForegroundColor Yellow
}

# ------------------------------------------------------------
# 5. LICENCIA DE WINDOWS
# ------------------------------------------------------------
Write-Host "5/6 Verificando licencia de Windows..." -ForegroundColor Cyan

$licOK = $false
$licObj = Get-CimInstance SoftwareLicensingProduct -ErrorAction SilentlyContinue |
          Where-Object { $_.Name -match "Windows" -and $_.PartialProductKey }

if ($licObj) {
    $licItem = $licObj | Select-Object -First 1

    $statusNum = 0
    if ($licItem.LicenseStatus -ne $null) {
        try { $statusNum = [int]"$($licItem.LicenseStatus)" } catch { $statusNum = 0 }
    }

    if     ($statusNum -eq 1) { $data["winLicencia"] = "Activa" }
    elseif ($statusNum -eq 2) { $data["winLicencia"] = "Gracia" }
    elseif ($statusNum -eq 3) { $data["winLicencia"] = "Notificacion" }
    elseif ($statusNum -eq 4) { $data["winLicencia"] = "ExtendedGrace" }
    elseif ($statusNum -eq 5) { $data["winLicencia"] = "InvalidGrace" }
    elseif ($statusNum -eq 6) { $data["winLicencia"] = "Sin licencia (tamper)" }
    else                       { $data["winLicencia"] = "Sin licencia" }

    $desc = ""
    try { $desc = "$($licItem.Description)" } catch {}
    if     ($desc -match "OEM")           { $data["winLicenciaCanal"] = "OEM" }
    elseif ($desc -match "VOLUME|MAK|KMS") { $data["winLicenciaCanal"] = "Volumen" }
    elseif ($desc -match "RETAIL")        { $data["winLicenciaCanal"] = "Retail" }
    else                                  { $data["winLicenciaCanal"] = $desc }

    $nombreLic = ""
    try { $nombreLic = "$($licItem.Name)" } catch {}
    $data["winLicenciaTipo"] = $nombreLic -replace "Windows ",""

    $lic_st = $data["winLicencia"]
    $lic_ch = $data["winLicenciaCanal"]
    Write-Host "OK Licencia: $lic_st - Canal: $lic_ch" -ForegroundColor Green
} else {
    $data["winLicencia"] = "No detectada"
    Write-Host "AVISO Licencia de Windows no detectada" -ForegroundColor Yellow
}

# ------------------------------------------------------------
# 6. PROGRAMAS INSTALADOS
# ------------------------------------------------------------
Write-Host "6/6 Recopilando programas instalados..." -ForegroundColor Cyan

$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$programasRaw = $regPaths | ForEach-Object { Get-ItemProperty $_ -ErrorAction SilentlyContinue }

if ($programasRaw) {
    $programasList = $programasRaw |
        Where-Object { $_.DisplayName -and $_.DisplayName.Trim() -ne "" -and -not $_.SystemComponent } |
        Sort-Object DisplayName -Unique |
        ForEach-Object {
            [ordered]@{
                nombre  = $_.DisplayName.Trim()
                editor  = if ($_.Publisher)      { $_.Publisher.Trim()      } else { "" }
                version = if ($_.DisplayVersion) { $_.DisplayVersion.Trim() } else { "" }
            }
        }
    $data["programas"] = @($programasList)
    $cnt = $data["programas"].Count
    Write-Host "OK $cnt programas encontrados" -ForegroundColor Green
} else {
    Write-Host "ERROR No se pudieron obtener programas" -ForegroundColor Red
}

# ------------------------------------------------------------
# EXPORTAR JSON - misma carpeta donde esta el script
# ------------------------------------------------------------
$hostname = $data["hostname"]

$scriptFolder = $PSScriptRoot
if (-not $scriptFolder) {
    $scriptFolder = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not $scriptFolder) {
    $scriptFolder = (Get-Location).Path
}

$outPath = Join-Path $scriptFolder "$hostname.json"

$jsonContent = $data | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($outPath, $jsonContent, [System.Text.UTF8Encoding]::new($false))

Write-Host ""
Write-Host "Archivo generado: $outPath" -ForegroundColor Green

Read-Host "Presiona ENTER para salir"
