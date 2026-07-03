<#
Exporta os dados de preco de energia da planilha "Curva Forward - CMF Energia.xlsx"
(abas Forward e Atac-I50-SECOeS) para um JSON consumido pelo prototipo do dashboard.

Reabre o Excel em uma instancia COM propria (invisivel, somente leitura) e sempre
fecha so essa instancia, sem tocar em janelas do Excel que o usuario ja tenha abertas.

Uso: abra um PowerShell nesta pasta e rode:
    .\export-data.ps1
#>

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$xlsxPath = Join-Path $root "Arqquivo Base\Curva Forward - CMF Energia.xlsx"
$outDir = Join-Path $root "data"
$outPath = Join-Path $outDir "precos-energia.json"

if (-not (Test-Path $xlsxPath)) {
    throw "Nao encontrei a planilha em: $xlsxPath"
}
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}

# Linhas 86 a 139 = 01/01/2022 a 01/06/2026, que e o intervalo com dados reais
# preenchidos em ambas as abas (confirmado via coluna Y, que so tem valor
# enquanto ha dado real; a partir da linha 140 as celulas ficam vazias porque
# sao meses futuros ainda nao alcancados).
$firstRow = 86
$lastRow = 139

# colunas K..T = anos-safra 2022..2031 (mesma ordem nas duas abas)
$vintageYears = 2022..2031
$vintageCols = @("K","L","M","N","O","P","Q","R","S","T")

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($xlsxPath, $false, $true)
    $wsF = $wb.Sheets.Item("Forward")
    $wsA = $wb.Sheets.Item("Atac-I50-SECOeS")

    function Get-Num($ws, $col, $row) {
        $v = $ws.Range("$col$row").Value2
        if ($null -eq $v -or $v -eq "") { return $null }
        return [math]::Round([double]$v, 2)
    }

    function Get-DateStr($ws, $row) {
        $v = $ws.Range("C$row").Value2
        $d = [datetime]::FromOADate($v)
        return $d.ToString("yyyy-MM")
    }

    $months = @()
    for ($r = $firstRow; $r -le $lastRow; $r++) {
        $months += (Get-DateStr $wsF $r)
    }

    # ano ainda em negociacao = tem valor na ultima linha (mes mais recente) da aba Forward
    $vintages = [ordered]@{}
    for ($i = 0; $i -lt $vintageYears.Count; $i++) {
        $year = $vintageYears[$i]
        $col = $vintageCols[$i]

        $nominal = @()
        $indexed = @()
        for ($r = $firstRow; $r -le $lastRow; $r++) {
            $nominal += (Get-Num $wsF $col $r)
            $indexed += (Get-Num $wsA $col $r)
        }

        $active = $null -ne (Get-Num $wsF $col $lastRow)

        $vintages["$year"] = [ordered]@{
            nominal = $nominal
            indexado = $indexed
            emNegociacao = $active
        }
    }

    $avg60mIndexado = @()
    $avg60mHistorico = @()
    # break-even nominal = aba Forward AD:AG  |  indexado (IPCA) = aba Atac-I50 AE:AH
    $beNomVerde = @();  $beNomAmarela = @();  $beNomVermelhaI = @();  $beNomVermelhaII = @()
    $beIdxVerde = @();  $beIdxAmarela = @();  $beIdxVermelhaI = @();  $beIdxVermelhaII = @()
    $pctEconomia = @()

    for ($r = $firstRow; $r -le $lastRow; $r++) {
        $avg60mIndexado += (Get-Num $wsA "Y" $r)
        $avg60mHistorico += (Get-Num $wsA "Z" $r)
        $beNomVerde     += (Get-Num $wsF "AD" $r)
        $beNomAmarela   += (Get-Num $wsF "AE" $r)
        $beNomVermelhaI += (Get-Num $wsF "AF" $r)
        $beNomVermelhaII+= (Get-Num $wsF "AG" $r)
        $beIdxVerde     += (Get-Num $wsA "AE" $r)
        $beIdxAmarela   += (Get-Num $wsA "AF" $r)
        $beIdxVermelhaI += (Get-Num $wsA "AG" $r)
        $beIdxVermelhaII+= (Get-Num $wsA "AH" $r)
        $pctEconomia += (Get-Num $wsA "AJ" $r)
    }

    $result = [ordered]@{
        geradoEm = (Get-Date).ToString("s")
        arquivoOrigem = "Arqquivo Base\Curva Forward - CMF Energia.xlsx"
        meses = $months
        anosSafra = $vintages
        mercadoLivre = [ordered]@{
            mediaIndexada60m = $avg60mIndexado
            mediaHistorica60m = $avg60mHistorico
        }
        breakEven = [ordered]@{
            nominal = [ordered]@{
                verde = $beNomVerde
                amarela = $beNomAmarela
                vermelhaI = $beNomVermelhaI
                vermelhaII = $beNomVermelhaII
            }
            indexado = [ordered]@{
                verde = $beIdxVerde
                amarela = $beIdxAmarela
                vermelhaI = $beIdxVermelhaI
                vermelhaII = $beIdxVermelhaII
            }
        }
        percentualEconomia = $pctEconomia
    }

    $json = $result | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($outPath, $json, [System.Text.Encoding]::UTF8)

    Write-Output "OK: dados exportados para $outPath"
    Write-Output "Meses: $($months.Count) ($($months[0]) a $($months[-1]))"
}
finally {
    if ($wb) { $wb.Close($false) }
    if ($excel) { $excel.Quit() }
    if ($wsA) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsA) | Out-Null }
    if ($wsF) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsF) | Out-Null }
    if ($wb) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
}
