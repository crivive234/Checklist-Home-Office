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
$data["programas"]           = @()

# ------------------------------------------------------------
# 1. HARDWARE
# ------------------------------------------------------------
Write-Host "1/5 Recopilando hardware..." -ForegroundColor Cyan

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
# 5. PROGRAMAS INSTALADOS
# ------------------------------------------------------------
Write-Host "5/5 Recopilando programas instalados..." -ForegroundColor Cyan

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
