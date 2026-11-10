<#
 Inventario-Software.ps1  (v2)
 Coleta aplicativos e versoes instalados em servidores Windows remotos (2008 ou superior)
 e exporta o resultado em CSV, HTML ou PDF.

 Novidades da v2:
   - Coleta em paralelo (varios servidores ao mesmo tempo, configuravel na tela)
   - Filtro para ocultar atualizacoes do Windows (KB), alternavel a qualquer momento
     (as exportacoes respeitam o filtro que estiver ativo na tela)

 Como funciona: consulta o registro remoto via WMI (StdRegProv), lendo as chaves Uninstall
 (64 e 32 bits). Nao precisa instalar nada nos servidores e nao depende do servico
 "Remote Registry".

 Requisitos:
   - Estacao de execucao: Windows com PowerShell 3.0+ (recomendado 5.1)
   - Usuario com permissao de administrador nos servidores
   - Porta 135 (RPC/DCOM) liberada ate os servidores (padrao do WMI)
   - PDF: Microsoft Edge ou Google Chrome instalado na estacao
#>

# ---------- Configuracoes ----------
$Delimitador    = ';'   # ';' abre direto no Excel pt-BR. Use ',' se preferir.
$TimeoutPortaMs = 3000  # tempo maximo para testar a porta 135 em cada servidor
$ParalelismoPadrao = 10 # servidores consultados ao mesmo tempo (ajustavel na tela)

# ---------- Garante modo STA (necessario para a interface grafica) ----------
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Start-Process powershell.exe -ArgumentList '-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',"`"$PSCommandPath`""
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:colsExport = @('Servidor','Aplicativo','Versao','Fabricante','DataInstalacao','Arquitetura','DataColeta')

# ---------- Tabela de dados ----------
$dt = New-Object System.Data.DataTable
foreach ($c in ($script:colsExport + 'Atualizacao')) { [void]$dt.Columns.Add($c) }   # 'Atualizacao' (S/N) e interna, nao exportada

function Test-Atualizacao($nome) {
    return ($nome -match '\bKB\d{6,8}\b') -or
           ($nome -match '^(Security Update|Hotfix|Update Rollup|Cumulative Update)\b') -or
           ($nome -match '^Update for (Windows|Microsoft)')
}

function Add-Linha($r) {
    $flag = if (Test-Atualizacao $r.A) { 'S' } else { 'N' }
    [void]$dt.Rows.Add($r.S, $r.A, $r.V, $r.F, $r.D, $r.Arq, $r.Dt, $flag)
}

# ---------- Worker executado em paralelo (1 runspace por servidor) ----------
$worker = {
    param($servidor, $cred, $timeoutMs, $sync)

    $HKLM  = [uint32]2147483650
    $agora = Get-Date -Format 'dd/MM/yyyy HH:mm'

    function Nova($a, $v, $f, $d, $arq) {
        New-Object PSObject -Property @{ S = $servidor; A = $a; V = $v; F = $f; D = $d; Arq = $arq; Dt = $agora }
    }
    function Test-Porta($h, $porta, $ms) {
        try {
            $tcp = New-Object Net.Sockets.TcpClient
            $ar  = $tcp.BeginConnect($h, $porta, $null, $null)
            $ok  = $ar.AsyncWaitHandle.WaitOne($ms, $false)
            if ($ok) { try { $tcp.EndConnect($ar) } catch { $ok = $false } }
            $tcp.Close()
            return $ok
        } catch { return $false }
    }

    if ($sync.cancelar) { return }
    $local = ($servidor -in @('.', 'localhost', '127.0.0.1', $env:COMPUTERNAME))

    if (-not $local -and -not (Test-Porta $servidor 135 $timeoutMs)) {
        Nova '[FALHA]' 'Porta 135 inacessivel ou servidor offline' '' '' ''
        return
    }

    try {
        $p = @{ List = $true; Namespace = 'root\default'; ComputerName = $servidor; ErrorAction = 'Stop' }
        if ($cred -and -not $local) { $p.Credential = $cred }
        $reg = Get-WmiObject @p | Where-Object { $_.Name -eq 'StdRegProv' }
        if (-not $reg) { throw 'Provedor de registro WMI (StdRegProv) nao encontrado no servidor.' }

        $caminhos = @(
            @{ Path = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall';             Arq = '64 bits' },
            @{ Path = 'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'; Arq = '32 bits' }
        )

        foreach ($cam in $caminhos) {
            $chaves = $reg.EnumKey($HKLM, $cam.Path).sNames
            if (-not $chaves) { continue }
            foreach ($k in $chaves) {
                if ($sync.cancelar) { return }
                $sub  = "$($cam.Path)\$k"
                $nome = $reg.GetStringValue($HKLM, $sub, 'DisplayName').sValue
                if ([string]::IsNullOrWhiteSpace($nome)) { continue }
                $ver  = $reg.GetStringValue($HKLM, $sub, 'DisplayVersion').sValue
                $fab  = $reg.GetStringValue($HKLM, $sub, 'Publisher').sValue
                $data = $reg.GetStringValue($HKLM, $sub, 'InstallDate').sValue
                if ($data -match '^\d{8}$') {
                    $data = $data.Substring(6,2) + '/' + $data.Substring(4,2) + '/' + $data.Substring(0,4)
                }
                Nova $nome.Trim() $ver $fab $data $cam.Arq
            }
        }
    } catch {
        Nova '[FALHA]' $_.Exception.Message.Trim() '' '' ''
    }
}

# ---------- Funcoes de exportacao (respeitam o filtro ativo na tela) ----------
function Exportar-CSV($arquivo) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(($script:colsExport -join $Delimitador))
    foreach ($row in $dt.DefaultView) {
        $linha = foreach ($c in $script:colsExport) {
            '"' + ([string]$row[$c]).Replace('"', '""') + '"'
        }
        [void]$sb.AppendLine(($linha -join $Delimitador))
    }
    [IO.File]::WriteAllText($arquivo, $sb.ToString(), (New-Object Text.UTF8Encoding($true)))
}

function Montar-HTML {
    $enc = { param($t) [System.Web.HttpUtility]::HtmlEncode([string]$t) }
    $linhas  = @($dt.DefaultView)
    $qtdServ = @($linhas | ForEach-Object { $_['Servidor'] } | Select-Object -Unique).Count
    $qtdApps = @($linhas | Where-Object { $_['Aplicativo'] -ne '[FALHA]' }).Count
    $obs = if ($chkKB.Checked) { ' &mdash; atualizações do Windows (KB) ocultadas' } else { '' }
    $h = New-Object System.Text.StringBuilder
    [void]$h.AppendLine(@"
<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8">
<title>Inventário de Software</title>
<style>
 @page { size: A4 landscape; margin: 10mm; }
 body { font-family: Segoe UI, Arial, sans-serif; font-size: 12px; color: #222; margin: 20px; }
 h1 { font-size: 20px; margin: 0 0 4px; } .sub { color: #666; margin-bottom: 12px; }
 input { padding: 6px; width: 320px; margin-bottom: 10px; }
 table { border-collapse: collapse; width: 100%; }
 th { background: #1f4e79; color: #fff; text-align: left; padding: 6px; }
 td { border-bottom: 1px solid #ddd; padding: 5px 6px; }
 tr:nth-child(even) td { background: #f5f8fb; }
 tr.falha td { background: #fde8e8; color: #a00; }
 @media print { input { display: none; } th { -webkit-print-color-adjust: exact; print-color-adjust: exact; } }
</style></head><body>
<h1>Inventário de Software</h1>
<div class="sub">Gerado em $(Get-Date -Format 'dd/MM/yyyy HH:mm') &mdash; $qtdServ servidor(es), $qtdApps aplicativo(s)$obs</div>
<input id="f" placeholder="Filtrar (servidor, aplicativo, versão...)" onkeyup="var q=this.value.toLowerCase();document.querySelectorAll('tbody tr').forEach(function(r){r.style.display=r.innerText.toLowerCase().indexOf(q)>-1?'':'none'})">
<table><thead><tr>
"@)
    foreach ($c in $script:colsExport) { [void]$h.Append("<th>$(& $enc $c)</th>") }
    [void]$h.AppendLine('</tr></thead><tbody>')
    foreach ($row in $linhas) {
        $cls = if ($row['Aplicativo'] -eq '[FALHA]') { ' class="falha"' } else { '' }
        [void]$h.Append("<tr$cls>")
        foreach ($c in $script:colsExport) { [void]$h.Append("<td>$(& $enc $row[$c])</td>") }
        [void]$h.AppendLine('</tr>')
    }
    [void]$h.AppendLine('</tbody></table></body></html>')
    return $h.ToString()
}

function Exportar-HTML($arquivo) {
    [IO.File]::WriteAllText($arquivo, (Montar-HTML), (New-Object Text.UTF8Encoding($true)))
}

function Exportar-PDF($arquivo) {
    $navegador = @(
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $navegador) { throw 'Microsoft Edge ou Google Chrome nao encontrado para gerar o PDF.' }

    $tmpHtml   = Join-Path $env:TEMP 'inventario_tmp.html'
    $tmpPerfil = Join-Path $env:TEMP 'inventario_perfil'
    [IO.File]::WriteAllText($tmpHtml, (Montar-HTML), (New-Object Text.UTF8Encoding($true)))
    if (Test-Path $arquivo) { Remove-Item $arquivo -Force }

    $url  = 'file:///' + ($tmpHtml -replace '\\', '/')
    $args = @('--headless', '--disable-gpu', '--no-pdf-header-footer',
              "--user-data-dir=`"$tmpPerfil`"", "--print-to-pdf=`"$arquivo`"", "`"$url`"")
    Start-Process -FilePath $navegador -ArgumentList $args -Wait -WindowStyle Hidden

    Remove-Item $tmpHtml -Force -ErrorAction SilentlyContinue
    Remove-Item $tmpPerfil -Recurse -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path $arquivo)) { throw 'Falha ao gerar o PDF.' }
}

function Salvar-Como($filtro, $ext, $acao) {
    if ($dt.DefaultView.Count -eq 0) {
        [void][Windows.Forms.MessageBox]::Show('Nao ha dados para exportar. Execute a coleta primeiro.', 'Aviso')
        return
    }
    $dlg = New-Object Windows.Forms.SaveFileDialog
    $dlg.Filter   = $filtro
    $dlg.FileName = 'inventario_software_' + (Get-Date -Format 'yyyyMMdd_HHmm') + $ext
    if ($dlg.ShowDialog() -eq 'OK') {
        try {
            $form.Cursor = 'WaitCursor'
            & $acao $dlg.FileName
            $form.Cursor = 'Default'
            if ([Windows.Forms.MessageBox]::Show("Arquivo salvo:`n$($dlg.FileName)`n`nDeseja abri-lo agora?", 'Concluido', 'YesNo', 'Information') -eq 'Yes') {
                Start-Process $dlg.FileName
            }
        } catch {
            $form.Cursor = 'Default'
            [void][Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Erro', 'OK', 'Error')
        }
    }
}

# ---------- Interface grafica ----------
$form = New-Object Windows.Forms.Form
$form.Text = 'Inventario de Software - Servidores Windows'
$form.ClientSize = New-Object Drawing.Size(1100, 665)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object Drawing.Size(900, 580)

$lbl = New-Object Windows.Forms.Label
$lbl.Text = 'Servidores (um por linha - nome ou IP):'
$lbl.Location = '10,10'; $lbl.AutoSize = $true

$txt = New-Object Windows.Forms.TextBox
$txt.Multiline = $true; $txt.ScrollBars = 'Vertical'
$txt.Location = '10,32'; $txt.Size = '280,350'; $txt.Anchor = 'Top,Bottom,Left'

$btnImp = New-Object Windows.Forms.Button
$btnImp.Text = 'Importar lista (.txt)'; $btnImp.Location = '10,392'; $btnImp.Size = '280,28'; $btnImp.Anchor = 'Bottom,Left'

$chk = New-Object Windows.Forms.CheckBox
$chk.Text = 'Usar credenciais alternativas'; $chk.Location = '10,428'; $chk.AutoSize = $true; $chk.Anchor = 'Bottom,Left'

$chkKB = New-Object Windows.Forms.CheckBox
$chkKB.Text = 'Ocultar atualizações do Windows (KB)'; $chkKB.Location = '10,452'; $chkKB.AutoSize = $true
$chkKB.Checked = $true; $chkKB.Anchor = 'Bottom,Left'

$lblPar = New-Object Windows.Forms.Label
$lblPar.Text = 'Servidores em paralelo:'; $lblPar.Location = '10,482'; $lblPar.AutoSize = $true; $lblPar.Anchor = 'Bottom,Left'

$numPar = New-Object Windows.Forms.NumericUpDown
$numPar.Minimum = 1; $numPar.Maximum = 50; $numPar.Value = $ParalelismoPadrao
$numPar.Location = '160,478'; $numPar.Size = '60,24'; $numPar.Anchor = 'Bottom,Left'

$btnRun = New-Object Windows.Forms.Button
$btnRun.Text = 'Coletar'; $btnRun.Location = '10,506'; $btnRun.Size = '135,34'; $btnRun.Anchor = 'Bottom,Left'
$btnRun.BackColor = [Drawing.Color]::FromArgb(31, 78, 121); $btnRun.ForeColor = 'White'; $btnRun.FlatStyle = 'Flat'

$btnCan = New-Object Windows.Forms.Button
$btnCan.Text = 'Cancelar'; $btnCan.Location = '155,506'; $btnCan.Size = '135,34'; $btnCan.Anchor = 'Bottom,Left'; $btnCan.Enabled = $false

$grid = New-Object Windows.Forms.DataGridView
$grid.Location = '300,32'; $grid.Size = '790,508'; $grid.Anchor = 'Top,Bottom,Left,Right'
$grid.DataSource = $dt; $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
$grid.AutoSizeColumnsMode = 'AllCells'; $grid.SelectionMode = 'FullRowSelect'
$grid.BackgroundColor = 'White'
$grid.Columns['Atualizacao'].Visible = $false

$btnCsv  = New-Object Windows.Forms.Button; $btnCsv.Text  = 'Exportar CSV';  $btnCsv.Location  = '300,552'; $btnCsv.Size  = '120,32'; $btnCsv.Anchor  = 'Bottom,Left'
$btnHtml = New-Object Windows.Forms.Button; $btnHtml.Text = 'Exportar HTML'; $btnHtml.Location = '430,552'; $btnHtml.Size = '120,32'; $btnHtml.Anchor = 'Bottom,Left'
$btnPdf  = New-Object Windows.Forms.Button; $btnPdf.Text  = 'Exportar PDF';  $btnPdf.Location  = '560,552'; $btnPdf.Size  = '120,32'; $btnPdf.Anchor  = 'Bottom,Left'

$bar = New-Object Windows.Forms.ProgressBar
$bar.Location = '10,598'; $bar.Size = '1080,20'; $bar.Anchor = 'Bottom,Left,Right'

$status = New-Object Windows.Forms.Label
$status.Text = 'Pronto.'; $status.Location = '10,626'; $status.Size = '1080,22'; $status.Anchor = 'Bottom,Left,Right'

$form.Controls.AddRange(@($lbl, $txt, $btnImp, $chk, $chkKB, $lblPar, $numPar, $btnRun, $btnCan, $grid, $btnCsv, $btnHtml, $btnPdf, $bar, $status))

# ---------- Filtro de atualizacoes (KB) ----------
function Aplicar-Filtro {
    $dt.DefaultView.RowFilter = if ($chkKB.Checked) { "Atualizacao = 'N'" } else { '' }
    $ocultos = $dt.Rows.Count - $dt.DefaultView.Count
    if ($dt.Rows.Count -gt 0) {
        $status.Text = "$($dt.DefaultView.Count) registro(s) exibido(s)" + $(if ($ocultos -gt 0) { " ($ocultos atualização(ões) oculta(s))" } else { '' })
    }
}
$chkKB.Add_CheckedChanged({ Aplicar-Filtro })
Aplicar-Filtro

# ---------- Eventos ----------
$sync = [hashtable]::Synchronized(@{ cancelar = $false })

$btnImp.Add_Click({
    $dlg = New-Object Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Texto/CSV (*.txt;*.csv)|*.txt;*.csv|Todos (*.*)|*.*'
    if ($dlg.ShowDialog() -eq 'OK') { $txt.Text = (Get-Content $dlg.FileName | Out-String).Trim() }
})

$btnCan.Add_Click({ $sync.cancelar = $true; $status.Text = 'Cancelando...' })

$btnRun.Add_Click({
    $servidores = @($txt.Lines | ForEach-Object { $_.Trim() } |
                    Where-Object { $_ -and -not $_.StartsWith('#') } | Select-Object -Unique)
    if ($servidores.Count -eq 0) {
        [void][Windows.Forms.MessageBox]::Show('Informe ao menos um servidor.', 'Aviso'); return
    }
    $cred = $null
    if ($chk.Checked) {
        $cred = Get-Credential -Message 'Credenciais com acesso administrativo aos servidores'
        if (-not $cred) { return }
    }

    $dt.Clear(); $sync.cancelar = $false
    $btnRun.Enabled = $false; $btnCan.Enabled = $true; $numPar.Enabled = $false
    $total = $servidores.Count
    $bar.Maximum = $total; $bar.Value = 0
    $feitos = 0

    # Pool de execucao paralela
    $pool = [runspacefactory]::CreateRunspacePool(1, [int]$numPar.Value)
    $pool.Open()
    $jobs = New-Object System.Collections.ArrayList
    foreach ($s in $servidores) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript($worker).AddArgument($s).AddArgument($cred).AddArgument($TimeoutPortaMs).AddArgument($sync)
        [void]$jobs.Add(@{ PS = $ps; Handle = $ps.BeginInvoke(); Servidor = $s })
    }

    $status.Text = "Coletando 0 de $total servidor(es) ($([int]$numPar.Value) em paralelo)..."
    while ($jobs.Count -gt 0) {
        foreach ($j in @($jobs)) {
            if ($j.Handle.IsCompleted) {
                try {
                    $res = $j.PS.EndInvoke($j.Handle)
                    foreach ($r in $res) { Add-Linha $r }
                } catch {
                    Add-Linha (New-Object PSObject -Property @{ S = $j.Servidor; A = '[FALHA]'; V = $_.Exception.Message; F = ''; D = ''; Arq = ''; Dt = (Get-Date -Format 'dd/MM/yyyy HH:mm') })
                }
                $j.PS.Dispose(); $jobs.Remove($j)
                $feitos++
                $bar.Value = $feitos
                if (-not $sync.cancelar) { $status.Text = "Coletando $feitos de $total servidor(es) ($([int]$numPar.Value) em paralelo)..." }
            }
        }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    }
    $pool.Close(); $pool.Dispose()

    $dt.DefaultView.Sort = 'Servidor, Aplicativo'
    $falhas = @($dt.Rows | Where-Object { $_['Aplicativo'] -eq '[FALHA]' }).Count
    $prefixo = if ($sync.cancelar) { 'Coleta cancelada. ' } else { "Concluído: $total servidor(es), $falhas com falha. " }
    Aplicar-Filtro
    $status.Text = $prefixo + $status.Text
    $btnRun.Enabled = $true; $btnCan.Enabled = $false; $numPar.Enabled = $true
})

$btnCsv.Add_Click({  Salvar-Como 'CSV (*.csv)|*.csv'   '.csv'  { param($f) Exportar-CSV  $f } })
$btnHtml.Add_Click({ Salvar-Como 'HTML (*.html)|*.html' '.html' { param($f) Exportar-HTML $f } })
$btnPdf.Add_Click({  Salvar-Como 'PDF (*.pdf)|*.pdf'   '.pdf'  { param($f) Exportar-PDF  $f } })

[void]$form.ShowDialog()
