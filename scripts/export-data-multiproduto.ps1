<#
Extrai TODOS os 12 produtos (Modalidade x Tipo x Submercado), cada um com preco
nominal e indexado (IPCA), da planilha "Curva Forward - CMF Energia.xlsx" para um
JSON multiproduto consumido pelo dashboard.

Fonte por produto:
  indexado (IPCA) = aba "<Mod>-<Tipo>-<Sub>"       (ex.: Varej-Conv-NE)
  nominal         = aba "<Mod>-<Tipo>-<Sub>-Nom"   (ex.: Varej-Conv-NE-Nom)

Mapeamento (confirmado por regressao contra data/precos-energia.json):
  linhas 86..139           = 54 meses (2022-01 .. 2026-06)
  colunas K..T             = safras 2022..2031 (1 coluna por ano)
  Y / Z                    = mercado livre: media futura 60m / media historica
  AE:AH                    = break-even 4 bandeiras (verde/amarela/vermelhaI/vermelhaII)
  AJ                       = percentual de economia

Somente leitura: copia o xlsx p/ o scratchpad, abre read-only numa instancia COM
propria (invisivel) e nunca escreve na planilha. Grava so o JSON de saida.

Uso:
    .\export-data-multiproduto.ps1 [-OutPath <arquivo.json>] [-Validate]
Sem -OutPath, grava em data\precos-energia-multiproduto.json.
#>
param(
    [string]$OutPath,
    [switch]$Validate
)
$ErrorActionPreference = "Stop"

$root     = Split-Path -Parent $PSScriptRoot
$xlsxPath = Join-Path $root "Arqquivo Base\Curva Forward - CMF Energia.xlsx"
$baseline = Join-Path $root "data\precos-energia.json"
if (-not $OutPath) { $OutPath = Join-Path $root "data\precos-energia-multiproduto.json" }

$scratch = $env:TEMP
$xlsxTmp = Join-Path $scratch "curva-forward-extract.xlsx"

if (-not (Test-Path $xlsxPath)) { throw "Nao encontrei a planilha: $xlsxPath" }
Copy-Item -LiteralPath $xlsxPath -Destination $xlsxTmp -Force

$firstRow = 86; $lastRow = 139
$vintageYears = 2022..2031
$vintageCols  = @("K","L","M","N","O","P","Q","R","S","T")

$submercados = @(
    [ordered]@{ id = "SECOeS"; nome = "Sudeste / Centro-Oeste / Sul" }
    [ordered]@{ id = "NE";     nome = "Nordeste" }
    [ordered]@{ id = "N";      nome = "Norte" }
)
$modalidades = @(
    [ordered]@{ id = "Atac";  nome = "Atacadista" }
    [ordered]@{ id = "Varej"; nome = "Varejista" }
)
$tipos = @(
    [ordered]@{ id = "I50";  nome = "Incentivada 50%" }
    [ordered]@{ id = "Conv"; nome = "Convencional" }
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false

function Get-Num($ws, $col, $row) {
    $v = $ws.Range("$col$row").Value2
    if ($null -eq $v -or $v -eq "") { return $null }
    return [math]::Round([double]$v, 2)
}
function Get-ColArr($ws, $col) {
    $out = @()
    for ($r = $firstRow; $r -le $lastRow; $r++) { $out += (Get-Num $ws $col $r) }
    return ,$out
}
function Get-DateStr($ws, $row) {
    $v = $ws.Range("C$row").Value2
    return ([datetime]::FromOADate($v)).ToString("yyyy-MM")
}
# Monta o objeto de um produto a partir das abas indexada e nominal ja abertas.
# Usado tanto para os 12 produtos base quanto para as variantes "-SemEnc" (varejista
# sem os encargos), que tem exatamente o mesmo layout de colunas.
function Build-Product($wsIdx, $wsNom, $modId, $tipoId, $subId) {
    $anos = [ordered]@{}
    for ($i = 0; $i -lt $vintageYears.Count; $i++) {
        $y = "$($vintageYears[$i])"; $c = $vintageCols[$i]
        $emNeg = $null -ne (Get-Num $wsNom $c $lastRow)
        $anos[$y] = [ordered]@{
            nominal      = (Get-ColArr $wsNom $c)
            indexado     = (Get-ColArr $wsIdx $c)
            emNegociacao = $emNeg
        }
    }
    return [ordered]@{
        modalidade = $modId
        tipo       = $tipoId
        submercado = $subId
        anosSafra  = $anos
        mercadoLivre = [ordered]@{
            indexado = [ordered]@{ media60m = (Get-ColArr $wsIdx "Y"); mediaHist = (Get-ColArr $wsIdx "Z") }
            nominal  = [ordered]@{ media60m = (Get-ColArr $wsNom "Y"); mediaHist = (Get-ColArr $wsNom "Z") }
        }
        breakEven = [ordered]@{
            nominal  = [ordered]@{ verde = (Get-ColArr $wsNom "AE"); amarela = (Get-ColArr $wsNom "AF"); vermelhaI = (Get-ColArr $wsNom "AG"); vermelhaII = (Get-ColArr $wsNom "AH") }
            indexado = [ordered]@{ verde = (Get-ColArr $wsIdx "AE"); amarela = (Get-ColArr $wsIdx "AF"); vermelhaI = (Get-ColArr $wsIdx "AG"); vermelhaII = (Get-ColArr $wsIdx "AH") }
        }
        percentualEconomia = (Get-ColArr $wsIdx "AJ")
    }
}

try {
    $wb = $excel.Workbooks.Open($xlsxTmp, $false, $true)

    $wsFwd = $wb.Sheets.Item("Forward")
    $months = @()
    for ($r = $firstRow; $r -le $lastRow; $r++) { $months += (Get-DateStr $wsFwd $r) }

    $produtos = [ordered]@{}
    $summary  = @()

    foreach ($mod in $modalidades) {
      foreach ($tipo in $tipos) {
        foreach ($sub in $submercados) {
            $key    = "$($mod.id)-$($tipo.id)-$($sub.id)"
            $wsIdx = $wb.Sheets.Item($key)
            $wsNom = $wb.Sheets.Item("$key-Nom")

            $produtos[$key] = Build-Product $wsIdx $wsNom $mod.id $tipo.id $sub.id

            $ativos = 0
            foreach ($yk in $produtos[$key].anosSafra.Keys) { if ($produtos[$key].anosSafra[$yk].emNegociacao) { $ativos++ } }
            $beVn = $produtos[$key].breakEven.nominal.verde[0]
            $beVi = $produtos[$key].breakEven.indexado.verde[0]
            $summary += ("  {0,-24} anos ativos: {1,2}   verde(nom/idx) linha0: {2}/{3}" -f $key, $ativos, $beVn, $beVi)

            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsIdx) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsNom) | Out-Null

            # Varejista tem a variante sem encargos (abas "-SemEnc" / "-SemEnc-Nom").
            # Mesmo layout de colunas, entao reusa Build-Product. Atacadista nao tem
            # encargo embutido, logo nao ha variante.
            if ($mod.id -eq "Varej") {
                $keySe   = "$key-SemEnc"
                $wsIdxSe = $wb.Sheets.Item($keySe)
                $wsNomSe = $wb.Sheets.Item("$key-SemEnc-Nom")

                $produtos[$keySe] = Build-Product $wsIdxSe $wsNomSe $mod.id $tipo.id $sub.id

                $ativosSe = 0
                foreach ($yk in $produtos[$keySe].anosSafra.Keys) { if ($produtos[$keySe].anosSafra[$yk].emNegociacao) { $ativosSe++ } }
                $summary += ("  {0,-24} anos ativos: {1,2}   (sem encargos)" -f $keySe, $ativosSe)

                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsIdxSe) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsNomSe) | Out-Null
            }
        }
      }
    }

    $result = [ordered]@{
        geradoEm      = (Get-Date).ToString("s")
        arquivoOrigem = "Arqquivo Base\Curva Forward - CMF Energia.xlsx"
        meses         = $months
        submercados   = $submercados
        modalidades   = $modalidades
        tipos         = $tipos
        produtos      = $produtos
    }

    $json = $result | ConvertTo-Json -Depth 8
    [System.IO.File]::WriteAllText($OutPath, $json, [System.Text.Encoding]::UTF8)

    Write-Output "OK: $($produtos.Count) produtos exportados para $OutPath"
    Write-Output "Meses: $($months.Count) ($($months[0]) a $($months[-1]))"
    Write-Output ""
    Write-Output "Resumo por produto:"
    $summary | ForEach-Object { Write-Output $_ }

    # ---- Regressao contra o baseline single-product (Atac-I50-SECOeS) ----
    if ($Validate) {
        Write-Output ""
        Write-Output "== Regressao Atac-I50-SECOeS vs data/precos-energia.json =="
        $b = Get-Content -LiteralPath $baseline -Raw | ConvertFrom-Json
        $p = $produtos["Atac-I50-SECOeS"]
        $tol = 0.011; $fails = 0; $checks = 0
        function CmpArr($label, $got, $exp) {
            $script:checks++
            $n = [Math]::Max($got.Count, $exp.Count); $bad = 0
            for ($i = 0; $i -lt $n; $i++) {
                $g = $got[$i]; $e = $exp[$i]
                if ($null -eq $g -and $null -eq $e) { continue }
                if ($null -eq $g -or $null -eq $e -or [math]::Abs([double]$g - [double]$e) -gt $tol) { $bad++ }
            }
            if ($bad -eq 0) { Write-Output "  PASS  $label" }
            else { $script:fails++; Write-Output "  FAIL  $label -> $bad/$n divergem" }
        }
        foreach ($y in $vintageYears) {
            $ys = "$y"
            CmpArr "indexado $ys" $p.anosSafra.$ys.indexado $b.anosSafra.$ys.indexado
            CmpArr "nominal  $ys" $p.anosSafra.$ys.nominal  $b.anosSafra.$ys.nominal
        }
        CmpArr "mercadoLivre.media60m (Y)"  $p.mercadoLivre.indexado.media60m  $b.mercadoLivre.mediaIndexada60m
        CmpArr "mercadoLivre.mediaHist (Z)" $p.mercadoLivre.indexado.mediaHist $b.mercadoLivre.mediaHistorica60m
        CmpArr "percentualEconomia (AJ)"    $p.percentualEconomia              $b.percentualEconomia
        foreach ($band in "verde","amarela","vermelhaI","vermelhaII") {
            CmpArr "breakEven.indexado.$band" $p.breakEven.indexado.$band $b.breakEven.indexado.$band
        }
        Write-Output ""
        if ($fails -eq 0) { Write-Output "REGRESSAO: TODOS OS $checks CAMPOS DE PRECO BATERAM (PASS)" }
        else { Write-Output "REGRESSAO: $fails de $checks FALHARAM" }

        # break-even nominal: mudanca esperada (baseline usou Forward AD:AG; agora -Nom AE:AH)
        Write-Output ""
        Write-Output "== Break-even NOMINAL: baseline (Forward AD:AG) vs novo (-Nom AE:AH) =="
        foreach ($band in "verde","amarela","vermelhaI","vermelhaII") {
            $old = $b.breakEven.nominal.$band[0]
            $new = $p.breakEven.nominal.$band[0]
            Write-Output ("  {0,-11} baseline {1,8}  ->  novo {2,8}  (delta {3})" -f $band, $old, $new, ([math]::Round($new-$old,2)))
        }
    }
}
finally {
    if ($wb) { $wb.Close($false) }
    if ($excel) { $excel.Quit() }
    foreach ($o in @($wsFwd,$wb,$excel)) { if ($o) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($o) | Out-Null } }
    Remove-Item -LiteralPath $xlsxTmp -Force -ErrorAction SilentlyContinue
}
