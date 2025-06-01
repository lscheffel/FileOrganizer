Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Security

# Configuração inicial
$configPath = "$PSScriptRoot\config.ini"
$script:isProcessingSelection = $false

function Load-Config {
    param($path)
    $templates = @{}
    $settings = @{}
    if (Test-Path $path) {
        $lines = Get-Content $path -Raw -Encoding UTF8
        $currentSection = $null
        foreach ($line in $lines -split "`r?`n") {
            if ($line -match '^\[(.+)\]$') {
                $currentSection = $matches[1].Trim()
                if ($currentSection -eq "Settings") {
                    $settings = @{}
                } else {
                    $templates[$currentSection] = @{}
                }
            } elseif ($line -match '^\s*(.+?)=(.*)$' -and $currentSection) {
                if ($currentSection -eq "Settings") {
                    $settings[$matches[1].Trim()] = $matches[2].Trim()
                } else {
                    $templates[$currentSection][$matches[1].Trim()] = $matches[2].Trim()
                }
            }
        }
    }
    return @{ Templates = $templates; Settings = $settings }
}

function Save-Config {
    param($path, $templates, $settings)
    $content = "[Settings]`r`n"
    foreach ($key in $settings.Keys) {
        $content += "$key=$($settings[$key])`r`n"
    }
    $content += "`r`n"
    foreach ($templateName in $templates.Keys) {
        $content += "[$templateName]`r`n"
        foreach ($key in $templates[$templateName].Keys) {
            $content += "$key=$($templates[$templateName][$key])`r`n"
        }
        $content += "`r`n"
    }
    Set-Content -Path $path -Value $content -Encoding UTF8
}

function Show-Message {
    param($text, $title = "Aviso")
    [System.Windows.Forms.MessageBox]::Show($text, $title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

function Get-FileHashMD5 {
    param($filePath)
    try {
        $md5 = [System.Security.Cryptography.MD5]::Create()
        $stream = [System.IO.File]::OpenRead($filePath)
        $hashBytes = $md5.ComputeHash($stream)
        $stream.Close()
        return [BitConverter]::ToString($hashBytes).Replace("-", "").ToLower()
    } catch {
        return $null
    }
}

# Criar Formulário
$form = New-Object System.Windows.Forms.Form
$form.Text = "Organizador de Arquivos v3.3"
$form.Size = New-Object System.Drawing.Size(1280,720)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.MaximizeBox = $true
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable

# Aplicar Tema
function Apply-Theme {
    param($theme)
    if ($theme -eq "Neon") {
        $form.BackColor = [System.Drawing.Color]::Black
        $form.ForeColor = [System.Drawing.Color]::FromArgb(0,255,0) # Verde neon
        foreach ($control in $form.Controls) {
            if ($control -is [System.Windows.Forms.Panel]) {
                $control.BackColor = [System.Drawing.Color]::FromArgb(51,51,51) # Chumbo
                $control.Controls | ForEach-Object {
                    if ($_ -is [System.Windows.Forms.Button]) {
                        $_.BackColor = [System.Drawing.Color]::FromArgb(51,51,51) # Chumbo
                        $_.ForeColor = [System.Drawing.Color]::FromArgb(0,255,0)
                        $_.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                        $_.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(51,51,51)
                        $_.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(128,128,128) # Cinza
                    } elseif ($_ -is [System.Windows.Forms.TextBox] -or $_ -is [System.Windows.Forms.ListBox] -or $_ -is [System.Windows.Forms.ComboBox]) {
                        $_.BackColor = [System.Drawing.Color]::Black
                        $_.ForeColor = [System.Drawing.Color]::FromArgb(0,255,0)
                    } elseif ($_ -is [System.Windows.Forms.CheckBox] -or $_ -is [System.Windows.Forms.Label]) {
                        $_.BackColor = [System.Drawing.Color]::Black
                        $_.ForeColor = [System.Drawing.Color]::FromArgb(0,255,0)
                    }
                }
            }
        }
    } else {
        $form.BackColor = [System.Drawing.SystemColors]::Control
        $form.ForeColor = [System.Drawing.SystemColors]::ControlText
        foreach ($control in $form.Controls) {
            if ($control -is [System.Windows.Forms.Panel]) {
                $control.BackColor = [System.Drawing.SystemColors]::Control
                $control.Controls | ForEach-Object {
                    if ($_ -is [System.Windows.Forms.Button]) {
                        $_.BackColor = [System.Drawing.SystemColors]::Control
                        $_.ForeColor = [System.Drawing.SystemColors]::ControlText
                        $_.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
                        $_.FlatAppearance.MouseOverBackColor = [System.Drawing.SystemColors]::ControlLight
                    } elseif ($_ -is [System.Windows.Forms.TextBox] -or $_ -is [System.Windows.Forms.ListBox] -or $_ -is [System.Windows.Forms.ComboBox]) {
                        $_.BackColor = [System.Drawing.Color]::White
                        $_.ForeColor = [System.Drawing.SystemColors]::ControlText
                    } elseif ($_ -is [System.Windows.Forms.CheckBox] -or $_ -is [System.Windows.Forms.Label]) {
                        $_.BackColor = [System.Drawing.SystemColors]::Control
                        $_.ForeColor = [System.Drawing.SystemColors]::ControlText
                    }
                }
            }
        }
    }
}

# Controles - Coluna Esquerda (Configurações)
$panelConfig = New-Object System.Windows.Forms.Panel
$panelConfig.Location = New-Object System.Drawing.Point(10,10)
$panelConfig.Size = New-Object System.Drawing.Size(320,620)
$panelConfig.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$panelConfig.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Bottom
$form.Controls.Add($panelConfig)

$labelOrigem = New-Object System.Windows.Forms.Label
$labelOrigem.Location = New-Object System.Drawing.Point(10,10)
$labelOrigem.Size = New-Object System.Drawing.Size(300,20)
$labelOrigem.Text = "Pastas de Origem:"
$panelConfig.Controls.Add($labelOrigem)

$listBoxOrigem = New-Object System.Windows.Forms.ListBox
$listBoxOrigem.Location = New-Object System.Drawing.Point(10,30)
$listBoxOrigem.Size = New-Object System.Drawing.Size(300,150)
$panelConfig.Controls.Add($listBoxOrigem)

$buttonAddOrigem = New-Object System.Windows.Forms.Button
$buttonAddOrigem.Location = New-Object System.Drawing.Point(10,190)
$buttonAddOrigem.Size = New-Object System.Drawing.Size(90,25)
$buttonAddOrigem.Text = "Adicionar"
$panelConfig.Controls.Add($buttonAddOrigem)

$buttonRemoveOrigem = New-Object System.Windows.Forms.Button
$buttonRemoveOrigem.Location = New-Object System.Drawing.Point(110,190)
$buttonRemoveOrigem.Size = New-Object System.Drawing.Size(90,25)
$buttonRemoveOrigem.Text = "Remover"
$panelConfig.Controls.Add($buttonRemoveOrigem)

$buttonClearForm = New-Object System.Windows.Forms.Button
$buttonClearForm.Location = New-Object System.Drawing.Point(210,190)
$buttonClearForm.Size = New-Object System.Drawing.Size(90,25)
$buttonClearForm.Text = "Limpar Tudo"
$panelConfig.Controls.Add($buttonClearForm)

$labelDestino = New-Object System.Windows.Forms.Label
$labelDestino.Location = New-Object System.Drawing.Point(10,230)
$labelDestino.Size = New-Object System.Drawing.Size(300,20)
$labelDestino.Text = "Pasta de Destino:"
$panelConfig.Controls.Add($labelDestino)

$textBoxDestino = New-Object System.Windows.Forms.TextBox
$textBoxDestino.Location = New-Object System.Drawing.Point(10,250)
$textBoxDestino.Size = New-Object System.Drawing.Size(300,20)
$panelConfig.Controls.Add($textBoxDestino)

$buttonSelecionarDestino = New-Object System.Windows.Forms.Button
$buttonSelecionarDestino.Location = New-Object System.Drawing.Point(10,280)
$buttonSelecionarDestino.Size = New-Object System.Drawing.Size(140,25)
$buttonSelecionarDestino.Text = "Selecionar"
$panelConfig.Controls.Add($buttonSelecionarDestino)

$buttonClearDestino = New-Object System.Windows.Forms.Button
$buttonClearDestino.Location = New-Object System.Drawing.Point(160,280)
$buttonClearDestino.Size = New-Object System.Drawing.Size(140,25)
$buttonClearDestino.Text = "Limpar"
$panelConfig.Controls.Add($buttonClearDestino)

$labelConfig = New-Object System.Windows.Forms.Label
$labelConfig.Location = New-Object System.Drawing.Point(10,320)
$labelConfig.Size = New-Object System.Drawing.Size(300,20)
$labelConfig.Text = "Configurações:"
$panelConfig.Controls.Add($labelConfig)

$checkBoxMover = New-Object System.Windows.Forms.CheckBox
$checkBoxMover.Location = New-Object System.Drawing.Point(10,340)
$checkBoxMover.Size = New-Object System.Drawing.Size(150,20)
$checkBoxMover.Text = "Mover arquivos"
$checkBoxMover.Checked = $false
$panelConfig.Controls.Add($checkBoxMover)

$checkBoxExcluirDuplicatas = New-Object System.Windows.Forms.CheckBox
$checkBoxExcluirDuplicatas.Location = New-Object System.Drawing.Point(160,340)
$checkBoxExcluirDuplicatas.Size = New-Object System.Drawing.Size(150,20)
$checkBoxExcluirDuplicatas.Text = "Excluir duplicatas"
$checkBoxExcluirDuplicatas.Checked = $true
$panelConfig.Controls.Add($checkBoxExcluirDuplicatas)

$checkBoxLixeira = New-Object System.Windows.Forms.CheckBox
$checkBoxLixeira.Location = New-Object System.Drawing.Point(10,360)
$checkBoxLixeira.Size = New-Object System.Drawing.Size(150,20)
$checkBoxLixeira.Text = "Usar lixeira"
$checkBoxLixeira.Checked = $true
$panelConfig.Controls.Add($checkBoxLixeira)

$checkBoxSubpastas = New-Object System.Windows.Forms.CheckBox
$checkBoxSubpastas.Location = New-Object System.Drawing.Point(160,360)
$checkBoxSubpastas.Size = New-Object System.Drawing.Size(150,20)
$checkBoxSubpastas.Text = "Subpastas"
$checkBoxSubpastas.Checked = $false
$panelConfig.Controls.Add($checkBoxSubpastas)

$checkBoxHash = New-Object System.Windows.Forms.CheckBox
$checkBoxHash.Location = New-Object System.Drawing.Point(10,380)
$checkBoxHash.Size = New-Object System.Drawing.Size(150,20)
$checkBoxHash.Text = "Comparar hash"
$checkBoxHash.Checked = $false
$panelConfig.Controls.Add($checkBoxHash)

$checkBoxAbrirDestino = New-Object System.Windows.Forms.CheckBox
$checkBoxAbrirDestino.Location = New-Object System.Drawing.Point(160,380)
$checkBoxAbrirDestino.Size = New-Object System.Drawing.Size(150,20)
$checkBoxAbrirDestino.Text = "Abrir destino"
$checkBoxAbrirDestino.Checked = $false
$panelConfig.Controls.Add($checkBoxAbrirDestino)

$checkBoxEncerrar = New-Object System.Windows.Forms.CheckBox
$checkBoxEncerrar.Location = New-Object System.Drawing.Point(10,400)
$checkBoxEncerrar.Size = New-Object System.Drawing.Size(150,20)
$checkBoxEncerrar.Text = "Encerrar após"
$checkBoxEncerrar.Checked = $false
$panelConfig.Controls.Add($checkBoxEncerrar)

$labelFiltro = New-Object System.Windows.Forms.Label
$labelFiltro.Location = New-Object System.Drawing.Point(10,430)
$labelFiltro.Size = New-Object System.Drawing.Size(100,20)
$labelFiltro.Text = "Filtro:"
$panelConfig.Controls.Add($labelFiltro)

$textBoxFiltro = New-Object System.Windows.Forms.TextBox
$textBoxFiltro.Location = New-Object System.Drawing.Point(110,430)
$textBoxFiltro.Size = New-Object System.Drawing.Size(190,20)
$textBoxFiltro.Text = "*.*"
$panelConfig.Controls.Add($textBoxFiltro)

$labelTema = New-Object System.Windows.Forms.Label
$labelTema.Location = New-Object System.Drawing.Point(10,460)
$labelTema.Size = New-Object System.Drawing.Size(100,20)
$labelTema.Text = "Tema:"
$panelConfig.Controls.Add($labelTema)

$comboBoxTema = New-Object System.Windows.Forms.ComboBox
$comboBoxTema.Location = New-Object System.Drawing.Point(110,460)
$comboBoxTema.Size = New-Object System.Drawing.Size(190,25)
$comboBoxTema.Items.AddRange(@("Claro","Neon"))
$comboBoxTema.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboBoxTema.SelectedIndex = 0
$panelConfig.Controls.Add($comboBoxTema)

$labelTemplates = New-Object System.Windows.Forms.Label
$labelTemplates.Location = New-Object System.Drawing.Point(10,490)
$labelTemplates.Size = New-Object System.Drawing.Size(300,20)
$labelTemplates.Text = "Templates:"
$panelConfig.Controls.Add($labelTemplates)

$comboBoxTemplates = New-Object System.Windows.Forms.ComboBox
$comboBoxTemplates.Location = New-Object System.Drawing.Point(10,510)
$comboBoxTemplates.Size = New-Object System.Drawing.Size(300,25)
$comboBoxTemplates.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$panelConfig.Controls.Add($comboBoxTemplates)

$labelTemplateName = New-Object System.Windows.Forms.Label
$labelTemplateName.Location = New-Object System.Drawing.Point(10,540)
$labelTemplateName.Size = New-Object System.Drawing.Size(100,20)
$labelTemplateName.Text = "Nome:"
$panelConfig.Controls.Add($labelTemplateName)

$textBoxTemplateName = New-Object System.Windows.Forms.TextBox
$textBoxTemplateName.Location = New-Object System.Drawing.Point(110,540)
$textBoxTemplateName.Size = New-Object System.Drawing.Size(190,20)
$panelConfig.Controls.Add($textBoxTemplateName)

$buttonSalvarTemplate = New-Object System.Windows.Forms.Button
$buttonSalvarTemplate.Location = New-Object System.Drawing.Point(10,570)
$buttonSalvarTemplate.Size = New-Object System.Drawing.Size(90,25)
$buttonSalvarTemplate.Text = "Salvar"
$panelConfig.Controls.Add($buttonSalvarTemplate)

$buttonEditarTemplate = New-Object System.Windows.Forms.Button
$buttonEditarTemplate.Location = New-Object System.Drawing.Point(110,570)
$buttonEditarTemplate.Size = New-Object System.Drawing.Size(90,25)
$buttonEditarTemplate.Text = "Editar"
$panelConfig.Controls.Add($buttonEditarTemplate)

$buttonExcluirTemplate = New-Object System.Windows.Forms.Button
$buttonExcluirTemplate.Location = New-Object System.Drawing.Point(210,570)
$buttonExcluirTemplate.Size = New-Object System.Drawing.Size(90,25)
$buttonExcluirTemplate.Text = "Excluir"
$panelConfig.Controls.Add($buttonExcluirTemplate)

# Controles - Coluna Direita (Visualização)
$panelVisualizacao = New-Object System.Windows.Forms.Panel
$panelVisualizacao.Location = New-Object System.Drawing.Point(340,10)
$panelVisualizacao.Size = New-Object System.Drawing.Size(920,620)
$panelVisualizacao.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$panelVisualizacao.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$form.Controls.Add($panelVisualizacao)

$labelLog = New-Object System.Windows.Forms.Label
$labelLog.Location = New-Object System.Drawing.Point(10,10)
$labelLog.Size = New-Object System.Drawing.Size(900,20)
$labelLog.Text = "Log de Operações:"
$panelVisualizacao.Controls.Add($labelLog)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(10,30)
$logBox.Size = New-Object System.Drawing.Size(900,300)
$logBox.Multiline = $true
$logBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$logBox.ReadOnly = $true
$logBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$panelVisualizacao.Controls.Add($logBox)

$labelPreview = New-Object System.Windows.Forms.Label
$labelPreview.Location = New-Object System.Drawing.Point(10,340)
$labelPreview.Size = New-Object System.Drawing.Size(900,20)
$labelPreview.Text = "Pré-visualização:"
$panelVisualizacao.Controls.Add($labelPreview)

$listBoxPreview = New-Object System.Windows.Forms.ListBox
$listBoxPreview.Location = New-Object System.Drawing.Point(10,360)
$listBoxPreview.Size = New-Object System.Drawing.Size(900,250)
$listBoxPreview.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$panelVisualizacao.Controls.Add($listBoxPreview)

# Rodapé (Ações)
$panelAcoes = New-Object System.Windows.Forms.Panel
$panelAcoes.Location = New-Object System.Drawing.Point(10,640)
$panelAcoes.Size = New-Object System.Drawing.Size(1250,50)
$panelAcoes.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$panelAcoes.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($panelAcoes)

$buttonPreview = New-Object System.Windows.Forms.Button
$buttonPreview.Location = New-Object System.Drawing.Point(10,10)
$buttonPreview.Size = New-Object System.Drawing.Size(150,30)
$buttonPreview.Text = "Pré-visualizar"
$panelAcoes.Controls.Add($buttonPreview)

$buttonExecutar = New-Object System.Windows.Forms.Button
$buttonExecutar.Location = New-Object System.Drawing.Point(170,10)
$buttonExecutar.Size = New-Object System.Drawing.Size(150,30)
$buttonExecutar.Text = "Executar"
$buttonExecutar.Enabled = $false
$panelAcoes.Controls.Add($buttonExecutar)

$buttonRestaurarLixeira = New-Object System.Windows.Forms.Button
$buttonRestaurarLixeira.Location = New-Object System.Drawing.Point(330,10)
$buttonRestaurarLixeira.Size = New-Object System.Drawing.Size(150,30)
$buttonRestaurarLixeira.Text = "Restaurar Lixeira"
$panelAcoes.Controls.Add($buttonRestaurarLixeira)

# Tema
$comboBoxTema.Add_SelectedIndexChanged({
    $config = Load-Config -path $configPath
    $config.Settings["Theme"] = $comboBoxTema.SelectedItem
    Save-Config -path $configPath -templates $config.Templates -settings $config.Settings
    Apply-Theme -theme $comboBoxTema.SelectedItem
})

# Carregar configurações
$config = Load-Config -path $configPath
if ($config.Settings["Theme"]) {
    $comboBoxTema.SelectedItem = $config.Settings["Theme"]
    Apply-Theme -theme $config.Settings["Theme"]
} else {
    Apply-Theme -theme "Claro"
}

# Funções de template
function Populate-TemplatesDropdown {
    $currentSelection = $comboBoxTemplates.SelectedItem
    $comboBoxTemplates.Items.Clear()
    $config = Load-Config -path $configPath
    $templates = $config.Templates
    foreach ($key in $templates.Keys) {
        $comboBoxTemplates.Items.Add($key) | Out-Null
    }
    if ($comboBoxTemplates.Items.Count -gt 0) {
        if ($currentSelection -and $comboBoxTemplates.Items.Contains($currentSelection)) {
            $comboBoxTemplates.SelectedItem = $currentSelection
        } else {
            $comboBoxTemplates.SelectedIndex = 0
        }
    }
    return $templates
}

$comboBoxTemplates.Add_SelectedIndexChanged({
    if ($script:isProcessingSelection) { return }
    $script:isProcessingSelection = $true
    try {
        $sel = $comboBoxTemplates.SelectedItem
        if ($sel) {
            $config = Load-Config -path $configPath
            $templates = $config.Templates
            if ($templates.ContainsKey($sel)) {
                $t = $templates[$sel]
                $listBoxOrigem.Items.Clear()
                if ($t.PastasOrigem) {
                    $t.PastasOrigem.Split(';') | ForEach-Object {
                        if ($_.Trim()) { $listBoxOrigem.Items.Add($_.Trim()) | Out-Null }
                    }
                }
                $textBoxDestino.Text = if ($t.PastaDestino) { $t.PastaDestino } else { "" }
                $checkBoxMover.Checked = if ($t.MoverArquivos) { [bool]::Parse($t.MoverArquivos) } else { $false }
                $checkBoxExcluirDuplicatas.Checked = if ($t.ExcluirDuplicatas) { [bool]::Parse($t.ExcluirDuplicatas) } else { $true }
                $checkBoxLixeira.Checked = if ($t.UsarLixeira) { [bool]::Parse($t.UsarLixeira) } else { $true }
                $checkBoxSubpastas.Checked = if ($t.UsarSubpastas) { [bool]::Parse($t.UsarSubpastas) } else { $false }
                $checkBoxHash.Checked = if ($t.UsarHash) { [bool]::Parse($t.UsarHash) } else { $false }
                $checkBoxAbrirDestino.Checked = if ($t.AbrirDestino) { [bool]::Parse($t.AbrirDestino) } else { $false }
                $checkBoxEncerrar.Checked = if ($t.EncerrarPrograma) { [bool]::Parse($t.EncerrarPrograma) } else { $false }
                $textBoxFiltro.Text = if ($t.Filtro) { $t.Filtro } else { "*.*" }
                $textBoxTemplateName.Text = $sel
            }
        }
    } catch {
        $logBox.AppendText("Erro ao carregar template: $_`r`n")
    } finally {
        $script:isProcessingSelection = $false
    }
})

$buttonSalvarTemplate.Add_Click({
    if (-not $textBoxTemplateName.Text) {
        Show-Message "Digite um nome para o template."
        return
    }
    $templateName = $textBoxTemplateName.Text.Trim()
    $config = Load-Config -path $configPath
    $templates = $config.Templates
    $templates[$templateName] = @{
        PastasOrigem = ($listBoxOrigem.Items -join ";")
        PastaDestino = $textBoxDestino.Text
        MoverArquivos = $checkBoxMover.Checked.ToString()
        ExcluirDuplicatas = $checkBoxExcluirDuplicatas.Checked.ToString()
        UsarLixeira = $checkBoxLixeira.Checked.ToString()
        UsarSubpastas = $checkBoxSubpastas.Checked.ToString()
        UsarHash = $checkBoxHash.Checked.ToString()
        AbrirDestino = $checkBoxAbrirDestino.Checked.ToString()
        EncerrarPrograma = $checkBoxEncerrar.Checked.ToString()
        Filtro = $textBoxFiltro.Text
    }
    Save-Config -path $configPath -templates $templates -settings $config.Settings
    $comboBoxTemplates.Items.Add($templateName) | Out-Null
    $comboBoxTemplates.SelectedItem = $templateName
    Show-Message "Template '$templateName' salvo com sucesso."
})

$buttonEditarTemplate.Add_Click({
    if (-not $comboBoxTemplates.SelectedItem) {
        Show-Message "Selecione um template para editar."
        return
    }
    if (-not $textBoxTemplateName.Text) {
        Show-Message "Digite um nome para o template."
        return
    }
    $templateName = $textBoxTemplateName.Text.Trim()
    $config = Load-Config -path $configPath
    $templates = $config.Templates
    $oldName = $comboBoxTemplates.SelectedItem
    $templates[$templateName] = @{
        PastasOrigem = ($listBoxOrigem.Items -join ";")
        PastaDestino = $textBoxDestino.Text
        MoverArquivos = $checkBoxMover.Checked.ToString()
        ExcluirDuplicatas = $checkBoxExcluirDuplicatas.Checked.ToString()
        UsarLixeira = $checkBoxLixeira.Checked.ToString()
        UsarSubpastas = $checkBoxSubpastas.Checked.ToString()
        UsarHash = $checkBoxHash.Checked.ToString()
        AbrirDestino = $checkBoxAbrirDestino.Checked.ToString()
        EncerrarPrograma = $checkBoxEncerrar.Checked.ToString()
        Filtro = $textBoxFiltro.Text
    }
    if ($templateName -ne $oldName) {
        $templates.Remove($oldName)
    }
    Save-Config -path $configPath -templates $templates -settings $config.Settings
    $comboBoxTemplates.Items.Clear()
    foreach ($key in $templates.Keys) {
        $comboBoxTemplates.Items.Add($key) | Out-Null
    }
    $comboBoxTemplates.SelectedItem = $templateName
    Show-Message "Template '$templateName' editado com sucesso."
})

$buttonExcluirTemplate.Add_Click({
    if (-not $comboBoxTemplates.SelectedItem) {
        Show-Message "Selecione um template para excluir."
        return
    }
    $templateName = $comboBoxTemplates.SelectedItem
    $confirm = [System.Windows.Forms.MessageBox]::Show("Deseja excluir o template '$templateName'?", "Confirmar Exclusão", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        $config = Load-Config -path $configPath
        $templates = $config.Templates
        $templates.Remove($templateName)
        Save-Config -path $configPath -templates $templates -settings $config.Settings
        $comboBoxTemplates.Items.Clear()
        foreach ($key in $templates.Keys) {
            $comboBoxTemplates.Items.Add($key) | Out-Null
        }
        if ($comboBoxTemplates.Items.Count -gt 0) {
            $comboBoxTemplates.SelectedIndex = 0
        }
        $textBoxTemplateName.Text = ""
        Show-Message "Template '$templateName' excluído com sucesso."
    }
})

# Adicionar/Remover pasta origem
$buttonAddOrigem.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Selecione uma pasta de origem"
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        if (-not $listBoxOrigem.Items.Contains($folderBrowser.SelectedPath)) {
            $listBoxOrigem.Items.Add($folderBrowser.SelectedPath) | Out-Null
            $logBox.AppendText("Pasta adicionada: $($folderBrowser.SelectedPath)`r`n")
        }
    }
})

$buttonRemoveOrigem.Add_Click({
    if ($listBoxOrigem.SelectedItem) {
        $logBox.AppendText("Pasta removida: $($listBoxOrigem.SelectedItem)`r`n")
        $listBoxOrigem.Items.Remove($listBoxOrigem.SelectedItem)
    }
})

# Selecionar/Limpar pasta destino
$buttonSelecionarDestino.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Selecione a pasta de destino"
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxDestino.Text = $folderBrowser.SelectedPath
        $logBox.AppendText("Destino selecionado: $($folderBrowser.SelectedPath)`r`n")
    }
})

$buttonClearDestino.Add_Click({
    $textBoxDestino.Text = ""
    $logBox.AppendText("Pasta de destino limpa`r`n")
})

# Limpar formulário
$buttonClearForm.Add_Click({
    $listBoxOrigem.Items.Clear()
    $textBoxDestino.Text = ""
    $checkBoxMover.Checked = $false
    $checkBoxExcluirDuplicatas.Checked = $true
    $checkBoxLixeira.Checked = $true
    $checkBoxSubpastas.Checked = $false
    $checkBoxHash.Checked = $false
    $checkBoxAbrirDestino.Checked = $false
    $checkBoxEncerrar.Checked = $false
    $textBoxFiltro.Text = "*.*"
    $textBoxTemplateName.Text = ""
    $listBoxPreview.Items.Clear()
    $buttonExecutar.Enabled = $false
    $logBox.AppendText("Formulário limpo`r`n")
})

# Restaurar da Lixeira
$buttonRestaurarLixeira.Add_Click({
    $shell = New-Object -ComObject Shell.Application
    $recycleBin = $shell.NameSpace(0xA)
    $items = $recycleBin.Items()
    if ($items.Count -eq 0) {
        Show-Message "A lixeira está vazia."
        return
    }
    $listForm = New-Object System.Windows.Forms.Form
    $listForm.Text = "Restaurar Arquivos"
    $listForm.Size = New-Object System.Drawing.Size(500,400)
    $listForm.StartPosition = "CenterScreen"
    $listForm.BackColor = [System.Drawing.Color]::Black
    $listForm.ForeColor = [System.Drawing.Color]::FromArgb(0,255,0)
    $listBoxRestore = New-Object System.Windows.Forms.ListBox
    $listBoxRestore.Location = New-Object System.Drawing.Point(10,10)
    $listBoxRestore.Size = New-Object System.Drawing.Size(460,300)
    $listBoxRestore.SelectionMode = "MultiExtended"
    $listBoxRestore.BackColor = [System.Drawing.Color]::Black
    $listBoxRestore.ForeColor = [System.Drawing.Color]::FromArgb(0,255,0)
    foreach ($item in $items) {
        $listBoxRestore.Items.Add("$($item.Name) ($($item.Path))") | Out-Null
    }
    $listForm.Controls.Add($listBoxRestore)
    $buttonRestore = New-Object System.Windows.Forms.Button
    $buttonRestore.Location = New-Object System.Drawing.Point(10,320)
    $buttonRestore.Size = New-Object System.Drawing.Size(100,25)
    $buttonRestore.Text = "Restaurar"
    $buttonRestore.BackColor = [System.Drawing.Color]::FromArgb(51,51,51)
    $buttonRestore.ForeColor = [System.Drawing.Color]::FromArgb(0,255,0)
    $buttonRestore.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $buttonRestore.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(51,51,51)
    $buttonRestore.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(128,128,128)
    $buttonRestore.Add_Click({
        foreach ($selected in $listBoxRestore.SelectedItems) {
            $itemName = ($selected -split " \(")[0]
            foreach ($item in $recycleBin.Items()) {
                if ($item.Name -eq $itemName) {
                    $originalPath = $item.ExtendedProperty("LocalizedResourceName")
                    if (-not $originalPath) { $originalPath = Join-Path $textBoxDestino.Text $item.Name }
                    $parentFolder = [System.IO.Path]::GetDirectoryName($originalPath)
                    if (-not (Test-Path $parentFolder)) { New-Item -Path $parentFolder -ItemType Directory -Force | Out-Null }
                    $item.InvokeVerb("Restore")
                    $logBox.AppendText("Restaurado: $($item.Name) para $originalPath`r`n")
                    break
                }
            }
        }
        $listForm.Close()
    })
    $listForm.Controls.Add($buttonRestore)
    $listForm.ShowDialog()
})

# Pré-visualização
$buttonPreview.Add_Click({
    $listBoxPreview.Items.Clear()
    $buttonExecutar.Enabled = $false
    if ($listBoxOrigem.Items.Count -eq 0) {
        Show-Message "Adicione pelo menos uma pasta de origem." "Aviso"
        return
    }
    if (-not $textBoxDestino.Text) {
        Show-Message "Selecione uma pasta de destino." "Aviso"
        return
    }
    foreach ($origem in $listBoxOrigem.Items) {
        if (-not (Test-Path $origem)) {
            $listBoxPreview.Items.Add("Erro: Pasta de origem não encontrada: $origem")
            continue
        }
        $files = Get-ChildItem -Path $origem -Filter $textBoxFiltro.Text -File -Recurse -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            $extension = [System.IO.Path]::GetExtension($file.Name).TrimStart('.').ToLower()
            $destFolder = if ($checkBoxSubpastas.Checked -and $extension) { Join-Path $textBoxDestino.Text $extension } else { $textBoxDestino.Text }
            $destFile = Join-Path -Path $destFolder -ChildPath $file.Name
            $action = if ($checkBoxMover.Checked) { "Mover" } else { "Copiar" }
            if (Test-Path $destFile) {
                if ($checkBoxExcluirDuplicatas.Checked) {
                    $isDuplicate = $true
                    if ($checkBoxHash.Checked) {
                        $srcHash = Get-FileHashMD5 -filePath $file.FullName
                        $destHash = Get-FileHashMD5 -filePath $destFile
                        $isDuplicate = $srcHash -and $destHash -and ($srcHash -eq $destHash)
                    }
                    if ($isDuplicate) {
                        $destino = if ($checkBoxLixeira.Checked) { "Lixeira" } else { "Permanentemente" }
                        $listBoxPreview.Items.Add("Excluir duplicata: $($file.FullName) -> $destino")
                    } else {
                        $listBoxPreview.Items.Add("${action} (renomear): $($file.FullName) -> $destFile")
                    }
                } else {
                    $listBoxPreview.Items.Add("${action} (renomear): $($file.FullName) -> $destFile")
                }
            } else {
                $listBoxPreview.Items.Add("${action}: $($file.FullName) -> $destFile")
            }
        }
    }
    if ($listBoxPreview.Items.Count -gt 0) {
        $buttonExecutar.Enabled = $true
        $logBox.AppendText("Pré-visualização gerada: $($listBoxPreview.Items.Count) ações`r`n")
    } else {
        $logBox.AppendText("Pré-visualização vazia: nenhum arquivo encontrado`r`n")
    }
})

# Processar arquivos
function Processar-Arquivos {
    param(
        [string[]]$PastasOrigem,
        [string]$PastaDestino,
        [string]$Filtro,
        [bool]$MoverArquivos,
        [bool]$ExcluirDuplicatas,
        [bool]$UsarLixeira,
        [bool]$UsarSubpastas,
        [bool]$UsarHash
    )
    $logBox.Clear()
    foreach ($origem in $PastasOrigem) {
        if (-not (Test-Path $origem)) {
            $logBox.AppendText("Pasta de origem não encontrada: $origem`r`n")
            continue
        }
        $files = Get-ChildItem -Path $origem -Filter $Filtro -File -Recurse -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            $extension = [System.IO.Path]::GetExtension($file.Name).TrimStart('.').ToLower()
            $destFolder = if ($UsarSubpastas -and $extension) { Join-Path $PastaDestino $extension } else { $PastaDestino }
            if (-not (Test-Path $destFolder)) { New-Item -Path $destFolder -ItemType Directory -Force | Out-Null }
            $destFile = Join-Path -Path $destFolder -ChildPath $file.Name
            if (Test-Path $destFile) {
                if ($ExcluirDuplicatas) {
                    $isDuplicate = $true
                    if ($UsarHash) {
                        $srcHash = Get-FileHashMD5 -filePath $file.FullName
                        $destHash = Get-FileHashMD5 -filePath $destFile
                        $isDuplicate = $srcHash -and $destHash -and ($srcHash -eq $destHash)
                    }
                    if ($isDuplicate) {
                        if ($UsarLixeira) {
                            try {
                                $shell = New-Object -ComObject Shell.Application
                                $folder = $shell.NameSpace(0xA)
                                $folder.MoveHere($file.FullName)
                                $logBox.AppendText("Duplicata movida para lixeira: $($file.FullName)`r`n")
                            } catch {
                                $lixeiraLocal = Join-Path $PastaDestino "Lixeira"
                                if (-not (Test-Path $lixeiraLocal)) { New-Item -Path $lixeiraLocal -ItemType Directory | Out-Null }
                                Move-Item -Path $file.FullName -Destination $lixeiraLocal -Force
                                $logBox.AppendText("Duplicata movida para lixeira local: $($file.FullName)`r`n")
                            }
                        } else {
                            Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
                            $logBox.AppendText("Duplicata excluída permanentemente: $($file.FullName)`r`n")
                        }
                    } else {
                        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                        $extension = [System.IO.Path]::GetExtension($file.Name)
                        $counter = 1
                        $newFileName = $file.Name
                        while (Test-Path (Join-Path $destFolder $newFileName)) {
                            $newFileName = "${baseName}_$counter$extension"
                            $counter++
                        }
                        $destFile = Join-Path -Path $destFolder -ChildPath $newFileName
                        if ($MoverArquivos) {
                            Move-Item -Path $file.FullName -Destination $destFile -Force
                            $logBox.AppendText("Movido (renomeado): $($file.FullName) -> $destFile`r")
                        } else {
                            Copy-Item -Path $file.FullName -Destination $destFile -Force
                            $logBox.AppendText("Copiado (renomeado): $($file.FullName) -> $destFile`r`n")
                        }
                    }
                } else {
                    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                    $extension = [System.IO.Path]::GetExtension($file.Name)
                    $counter = 1
                    $newFileName = $file.Name
                    while (Test-Path (Join-Path $destFolder $newFileName)) {
                        $newFileName = "${baseName}_$counter$extension"
                        $counter++
                    }
                    $destFile = Join-Path -Path $destFolder -ChildPath $newFileName
                    if ($MoverArquivos) {
                        Move-Item -Path $file.FullName -Destination $destFile -Force
                        $logBox.AppendText("Movido (renomeado): $($file.FullName) -> $destFile`r")
                    } else {
                        Copy-Item -Path $file.FullName -Destination $destFile -Force
                        $logBox.AppendText("Copiado (renomeado): $($file.FullName) -> $destFile`r`n")
                    }
                }
            } else {
                if ($MoverArquivos) {
                    Move-Item -Path $file.FullName -Destination $destFile -Force
                    $logBox.AppendText("Movido: $($file.FullName) -> $destFile`r")
                } else {
                    Copy-Item -Path $file.FullName -Destination $destFile -Force
                    $logBox.AppendText("Copiado: $($file.FullName) -> $destFile`r`n")
                }
            }
        }
    }
    $logBox.AppendText("Processamento concluído.`r`n")
}

# Botão Executar
$buttonExecutar.Add_Click({
    if ($listBoxOrigem.Items.Count -eq 0) {
        Show-Message "Adicione pelo menos uma pasta de origem." "Aviso"
        return
    }
    if (-not (Test-Path -Path $textBoxDestino.Text)) {
        try {
            New-Item -ItemType Directory -Path $textBoxDestino.Text -ErrorAction Stop | Out-Null
            $logBox.AppendText("Pasta destino criada: $($textBoxDestino.Text)`r`n")
        } catch {
            Show-Message "Não foi possível criar a pasta destino: $($textBoxDestino.Text)."
            return
        }
    }
    $buttonExecutar.Enabled = $false
    try {
        Processar-Arquivos -PastasOrigem $listBoxOrigem.Items -PastaDestino $textBoxDestino.Text -Filtro $textBoxFiltro.Text `
            -MoverArquivos $checkBoxMover.Checked -ExcluirDuplicatas $checkBoxExcluirDuplicatas.Checked -UsarLixeira $checkBoxLixeira.Checked `
            -UsarSubpastas $checkBoxSubpastas.Checked -UsarHash $checkBoxHash.Checked
        if ($checkBoxAbrirDestino.Checked) {
            Start-Process -FilePath "explorer.exe" -ArgumentList $textBoxDestino.Text
        }
        if ($checkBoxEncerrar.Checked) {
            $form.Close()
        }
    } catch {
        Show-Message "Erro: $_"
    } finally {
        $buttonExecutar.Enabled = $false
        $listBoxPreview.Items.Clear()
    }
})

# Inicializar templates
Populate-TemplatesDropdown

# Garantir foco na janela
$form.Add_Activated({ $form.Activate() })
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()