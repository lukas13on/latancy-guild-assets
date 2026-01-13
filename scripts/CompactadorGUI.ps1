# Script: CompactadorGUI.ps1
# Descrição: Compactador com interface gráfica completa e log pieces.csv limpo

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Configurações
$script:tamanhoParteMB = 90
$script:tamanhoParteBytes = $script:tamanhoParteMB * 1MB
$script:operacaoEmAndamento = $false

# Função para criar e mostrar o menu principal
function Show-MainMenu {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Compactador de Arquivos em Partes"
    $form.Size = New-Object System.Drawing.Size(500, 400)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    
    # Título
    $labelTitulo = New-Object System.Windows.Forms.Label
    $labelTitulo.Text = "COMPACTADOR DE ARQUIVOS EM PARTES"
    $labelTitulo.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $labelTitulo.ForeColor = [System.Drawing.Color]::DarkCyan
    $labelTitulo.Size = New-Object System.Drawing.Size(450, 40)
    $labelTitulo.Location = New-Object System.Drawing.Point(25, 20)
    $labelTitulo.TextAlign = "MiddleCenter"
    $form.Controls.Add($labelTitulo)
    
    # Subtítulo
    $labelSubtitulo = New-Object System.Windows.Forms.Label
    $labelSubtitulo.Text = "Tamanho atual das partes: $($script:tamanhoParteMB) MB"
    $labelSubtitulo.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $labelSubtitulo.ForeColor = [System.Drawing.Color]::Gray
    $labelSubtitulo.Size = New-Object System.Drawing.Size(450, 30)
    $labelSubtitulo.Location = New-Object System.Drawing.Point(25, 60)
    $labelSubtitulo.TextAlign = "MiddleCenter"
    $form.Controls.Add($labelSubtitulo)
    
    # Separador
    $separador = New-Object System.Windows.Forms.Label
    $separador.Text = "─" * 60
    $separador.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $separador.ForeColor = [System.Drawing.Color]::LightGray
    $separador.Size = New-Object System.Drawing.Size(450, 20)
    $separador.Location = New-Object System.Drawing.Point(25, 90)
    $separador.TextAlign = "MiddleCenter"
    $form.Controls.Add($separador)
    
    # Botão Compactar
    $btnCompactar = New-Object System.Windows.Forms.Button
    $btnCompactar.Text = "1. COMPACTAR ARQUIVO/PASTA"
    $btnCompactar.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnCompactar.ForeColor = [System.Drawing.Color]::White
    $btnCompactar.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $btnCompactar.Size = New-Object System.Drawing.Size(400, 50)
    $btnCompactar.Location = New-Object System.Drawing.Point(50, 130)
    $btnCompactar.FlatStyle = "Flat"
    $btnCompactar.FlatAppearance.BorderSize = 0
    $btnCompactar.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnCompactar.Add_Click({
        $form.Hide()
        Show-CompactarMenu
        $form.Close()
    })
    $form.Controls.Add($btnCompactar)
    
    # Botão Descompactar
    $btnDescompactar = New-Object System.Windows.Forms.Button
    $btnDescompactar.Text = "2. DESCOMPACTAR PARTES"
    $btnDescompactar.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnDescompactar.ForeColor = [System.Drawing.Color]::White
    $btnDescompactar.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $btnDescompactar.Size = New-Object System.Drawing.Size(400, 50)
    $btnDescompactar.Location = New-Object System.Drawing.Point(50, 190)
    $btnDescompactar.FlatStyle = "Flat"
    $btnDescompactar.FlatAppearance.BorderSize = 0
    $btnDescompactar.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnDescompactar.Add_Click({
        $form.Hide()
        Show-DescompactarMenu
        $form.Close()
    })
    $form.Controls.Add($btnDescompactar)
    
    # Botão Configurar Tamanho
    $btnConfigurar = New-Object System.Windows.Forms.Button
    $btnConfigurar.Text = "3. CONFIGURAR TAMANHO DAS PARTES"
    $btnConfigurar.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnConfigurar.ForeColor = [System.Drawing.Color]::White
    $btnConfigurar.BackColor = [System.Drawing.Color]::FromArgb(255, 193, 7)
    $btnConfigurar.Size = New-Object System.Drawing.Size(400, 50)
    $btnConfigurar.Location = New-Object System.Drawing.Point(50, 250)
    $btnConfigurar.FlatStyle = "Flat"
    $btnConfigurar.FlatAppearance.BorderSize = 0
    $btnConfigurar.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnConfigurar.Add_Click({
        $form.Hide()
        Show-ConfigurarTamanho
        $form.Close()
    })
    $form.Controls.Add($btnConfigurar)
    
    # Botão Sair
    $btnSair = New-Object System.Windows.Forms.Button
    $btnSair.Text = "4. SAIR"
    $btnSair.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnSair.ForeColor = [System.Drawing.Color]::White
    $btnSair.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $btnSair.Size = New-Object System.Drawing.Size(400, 50)
    $btnSair.Location = New-Object System.Drawing.Point(50, 310)
    $btnSair.FlatStyle = "Flat"
    $btnSair.FlatAppearance.BorderSize = 0
    $btnSair.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnSair.Add_Click({
        if (Confirmar-Saida) {
            $form.Close()
        }
    })
    $form.Controls.Add($btnSair)
    
    # Evento de fechamento do formulário
    $form.Add_FormClosing({
        param($sender, $e)
        if ($script:operacaoEmAndamento) {
            [System.Windows.Forms.MessageBox]::Show(
                "Não é possível sair enquanto uma operação está em andamento!",
                "Aviso",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            $e.Cancel = $true
        } elseif (-not $e.Cancel) {
            if (-not (Confirmar-Saida)) {
                $e.Cancel = $true
            }
        }
    })
    
    $form.ShowDialog() | Out-Null
}

# Função para confirmar saída
function Confirmar-Saida {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Deseja realmente sair do Compactador?",
        "Confirmar Saída",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    return ($result -eq [System.Windows.Forms.DialogResult]::Yes)
}

# Função para mostrar menu de compactação
function Show-CompactarMenu {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Compactar Arquivo/Pasta"
    $form.Size = New-Object System.Drawing.Size(500, 350)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    
    # Título
    $labelTitulo = New-Object System.Windows.Forms.Label
    $labelTitulo.Text = "SELECIONE O QUE DESEJA COMPACTAR"
    $labelTitulo.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $labelTitulo.ForeColor = [System.Drawing.Color]::DarkCyan
    $labelTitulo.Size = New-Object System.Drawing.Size(450, 40)
    $labelTitulo.Location = New-Object System.Drawing.Point(25, 20)
    $labelTitulo.TextAlign = "MiddleCenter"
    $form.Controls.Add($labelTitulo)
    
    # Informações
    $labelInfo = New-Object System.Windows.Forms.Label
    $labelInfo.Text = "Tamanho das partes: $($script:tamanhoParteMB) MB"
    $labelInfo.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $labelInfo.ForeColor = [System.Drawing.Color]::Gray
    $labelInfo.Size = New-Object System.Drawing.Size(450, 30)
    $labelInfo.Location = New-Object System.Drawing.Point(25, 60)
    $labelInfo.TextAlign = "MiddleCenter"
    $form.Controls.Add($labelInfo)
    
    # Botão Selecionar Arquivo
    $btnArquivo = New-Object System.Windows.Forms.Button
    $btnArquivo.Text = "SELECIONAR ARQUIVO ÚNICO"
    $btnArquivo.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnArquivo.ForeColor = [System.Drawing.Color]::White
    $btnArquivo.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $btnArquivo.Size = New-Object System.Drawing.Size(400, 60)
    $btnArquivo.Location = New-Object System.Drawing.Point(50, 110)
    $btnArquivo.FlatStyle = "Flat"
    $btnArquivo.FlatAppearance.BorderSize = 0
    $btnArquivo.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnArquivo.Add_Click({
        $form.Hide()
        $caminho = Selecionar-Arquivo
        if ($caminho) {
            $form.Close()
            Confirmar-Compactacao -caminho $caminho -tipo "arquivo"
        } else {
            $form.Show()
        }
    })
    $form.Controls.Add($btnArquivo)
    
    # Botão Selecionar Pasta
    $btnPasta = New-Object System.Windows.Forms.Button
    $btnPasta.Text = "SELECIONAR PASTA (TODOS OS ARQUIVOS)"
    $btnPasta.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnPasta.ForeColor = [System.Drawing.Color]::White
    $btnPasta.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $btnPasta.Size = New-Object System.Drawing.Size(400, 60)
    $btnPasta.Location = New-Object System.Drawing.Point(50, 180)
    $btnPasta.FlatStyle = "Flat"
    $btnPasta.FlatAppearance.BorderSize = 0
    $btnPasta.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnPasta.Add_Click({
        $form.Hide()
        $caminho = Selecionar-Pasta
        if ($caminho) {
            $form.Close()
            Confirmar-Compactacao -caminho $caminho -tipo "pasta"
        } else {
            $form.Show()
        }
    })
    $form.Controls.Add($btnPasta)
    
    # Botão Voltar
    $btnVoltar = New-Object System.Windows.Forms.Button
    $btnVoltar.Text = "VOLTAR AO MENU PRINCIPAL"
    $btnVoltar.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnVoltar.ForeColor = [System.Drawing.Color]::White
    $btnVoltar.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
    $btnVoltar.Size = New-Object System.Drawing.Size(400, 40)
    $btnVoltar.Location = New-Object System.Drawing.Point(50, 260)
    $btnVoltar.FlatStyle = "Flat"
    $btnVoltar.FlatAppearance.BorderSize = 0
    $btnVoltar.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnVoltar.Add_Click({
        $form.Hide()
        Show-MainMenu
        $form.Close()
    })
    $form.Controls.Add($btnVoltar)
    
    $form.ShowDialog() | Out-Null
}

# Função para selecionar arquivo
function Selecionar-Arquivo {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Selecione o arquivo para compactar"
    $dialog.Multiselect = $false
    $dialog.CheckFileExists = $true
    $dialog.Filter = "Todos os arquivos (*.*)|*.*"
    
    if ($dialog.ShowDialog() -eq 'OK') {
        return $dialog.FileName
    }
    return $null
}

# Função para selecionar pasta
function Selecionar-Pasta {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Selecione a pasta para compactar"
    $dialog.ShowNewFolderButton = $false
    
    if ($dialog.ShowDialog() -eq 'OK') {
        return $dialog.SelectedPath
    }
    return $null
}

# Função para confirmar compactação
function Confirmar-Compactacao {
    param(
        [string]$caminho,
        [string]$tipo
    )
    
    $itemName = Split-Path $caminho -Leaf
    $itemDir = Split-Path $caminho -Parent
    
    $mensagem = @"
RESUMO DA OPERAÇÃO:

Item: $itemName ($tipo)
Local: $itemDir
Tamanho por parte: $($script:tamanhoParteMB) MB

Deseja continuar com a compactação?
"@
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        $mensagem,
        "Confirmar Compactação",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Executar-Compactacao -caminho $caminho -tipo $tipo
    } else {
        Show-MainMenu
    }
}

# Função para executar a compactação
function Executar-Compactacao {
    param(
        [string]$caminho,
        [string]$tipo
    )
    
    $script:operacaoEmAndamento = $true
    
    try {
        # Mostrar tela de progresso
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Compactando..."
        $progressForm.Size = New-Object System.Drawing.Size(500, 200)
        $progressForm.StartPosition = "CenterScreen"
        $progressForm.FormBorderStyle = "FixedDialog"
        $progressForm.MaximizeBox = $false
        $progressForm.Topmost = $true
        
        # Label de status
        $labelStatus = New-Object System.Windows.Forms.Label
        $labelStatus.Text = "Preparando para compactar..."
        $labelStatus.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $labelStatus.Size = New-Object System.Drawing.Size(460, 30)
        $labelStatus.Location = New-Object System.Drawing.Point(20, 20)
        $progressForm.Controls.Add($labelStatus)
        
        # Barra de progresso
        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Size = New-Object System.Drawing.Size(460, 30)
        $progressBar.Location = New-Object System.Drawing.Point(20, 60)
        $progressBar.Style = "Marquee"
        $progressBar.MarqueeAnimationSpeed = 50
        $progressForm.Controls.Add($progressBar)
        
        # Label de detalhes
        $labelDetalhes = New-Object System.Windows.Forms.Label
        $labelDetalhes.Text = ""
        $labelDetalhes.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $labelDetalhes.ForeColor = [System.Drawing.Color]::Gray
        $labelDetalhes.Size = New-Object System.Drawing.Size(460, 50)
        $labelDetalhes.Location = New-Object System.Drawing.Point(20, 100)
        $progressForm.Controls.Add($labelDetalhes)
        
        # Mostrar formulário imediatamente
        $progressForm.Show()
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            # Criar diretório para partes
            $sourceDir = Split-Path $caminho -Parent
            $partsDir = Join-Path $sourceDir "parts"
            
            $labelStatus.Text = "Criando pasta para partes..."
            [System.Windows.Forms.Application]::DoEvents()
            
            if (Test-Path $partsDir) {
                Remove-Item "$partsDir\*" -Force -Recurse -ErrorAction SilentlyContinue
            } else {
                New-Item -ItemType Directory -Path $partsDir -Force | Out-Null
            }
            
            # Nome do arquivo ZIP
            $itemName = Split-Path $caminho -Leaf
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $zipFileName = "$itemName`_$timestamp.zip"
            $tempZipPath = Join-Path $env:TEMP $zipFileName
            $partsPrefix = Join-Path $partsDir "$itemName`_$timestamp.zip.part"
            
            # Compactar
            $labelStatus.Text = "Compactando $tipo..."
            $labelDetalhes.Text = "Isso pode levar alguns minutos..."
            [System.Windows.Forms.Application]::DoEvents()
            
            if ($tipo -eq "pasta") {
                Compress-Archive -Path "$caminho\*" -DestinationPath $tempZipPath -CompressionLevel Optimal
            } else {
                $tempDir = Join-Path $env:TEMP "temp_zip_$timestamp"
                New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
                Copy-Item $caminho $tempDir -Force
                Compress-Archive -Path "$tempDir\*" -DestinationPath $tempZipPath -CompressionLevel Optimal
                Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            }
            
            # Verificar se o arquivo foi criado
            if (-not (Test-Path $tempZipPath)) {
                throw "Falha ao criar arquivo ZIP"
            }
            
            # Dividir em partes
            $labelStatus.Text = "Dividindo em partes de $($script:tamanhoParteMB) MB..."
            $labelDetalhes.Text = "Criando arquivos .part..."
            [System.Windows.Forms.Application]::DoEvents()
            
            $fileInfo = Get-Item $tempZipPath
            $totalBytes = $fileInfo.Length
            $stream = [System.IO.File]::OpenRead($tempZipPath)
            $buffer = New-Object byte[] $script:tamanhoParteBytes
            $partNumber = 1
            $totalParts = [Math]::Ceiling($totalBytes / $script:tamanhoParteBytes)
            
            # Inicializar lista para o CSV
            $listaArquivos = @()

            try {
                while (($bytesRead = $stream.Read($buffer, 0, $script:tamanhoParteBytes)) -gt 0) {
                    $labelDetalhes.Text = "Criando parte $partNumber de $totalParts..."
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    $partPath = "$partsPrefix$partNumber"
                    $outStream = [System.IO.File]::OpenWrite($partPath)
                    $outStream.Write($buffer, 0, $bytesRead)
                    $outStream.Close()
                    
                    # Adicionar nome do arquivo gerado à lista como string pura
                    $nomeParte = Split-Path $partPath -Leaf
                    $listaArquivos += $nomeParte

                    $partNumber++
                }
            } finally {
                $stream.Close()
            }
            
            Remove-Item $tempZipPath -Force -ErrorAction SilentlyContinue
            
            # GERAR ARQUIVO CSV (Sem cabeçalho, sem aspas)
            $labelStatus.Text = "Gerando arquivo pieces.csv..."
            $csvPath = Join-Path $sourceDir "pieces.csv"
            
            # Set-Content escreve o array linha por linha, sem aspas e sem cabeçalho
            $listaArquivos | Set-Content -Path $csvPath -Encoding UTF8
            
            # Fechar formulário de progresso
            $progressForm.Close()
            
            # Mostrar resultados
            $mensagemResultado = @"
COMPACTAÇÃO CONCLUÍDA COM SUCESSO!

Partes criadas: $(($partNumber - 1)) arquivos
Tamanho total: $([Math]::Round($totalBytes / 1MB, 2)) MB
Local das partes: $partsDir
Lista gerada em: $csvPath

Deseja abrir a pasta das partes?
"@
            
            $result = [System.Windows.Forms.MessageBox]::Show(
                $mensagemResultado,
                "Concluído",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Process "explorer.exe" -ArgumentList $partsDir
            }
            
            Show-MainMenu
            
        } catch {
            $progressForm.Close()
            throw $_
        }
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Erro durante a compactação: $($_.Exception.Message)",
            "Erro na Compactação",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        Show-MainMenu
    } finally {
        $script:operacaoEmAndamento = $false
    }
}

# Função para mostrar menu de descompactação
function Show-DescompactarMenu {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Descompactar Partes"
    $form.Size = New-Object System.Drawing.Size(500, 300)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    
    # Título
    $labelTitulo = New-Object System.Windows.Forms.Label
    $labelTitulo.Text = "SELECIONE A PASTA COM AS PARTES"
    $labelTitulo.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $labelTitulo.ForeColor = [System.Drawing.Color]::DarkGreen
    $labelTitulo.Size = New-Object System.Drawing.Size(450, 40)
    $labelTitulo.Location = New-Object System.Drawing.Point(25, 20)
    $labelTitulo.TextAlign = "MiddleCenter"
    $form.Controls.Add($labelTitulo)
    
    # Instruções
    $labelInstrucoes = New-Object System.Windows.Forms.Label
    $labelInstrucoes.Text = "Selecione a pasta que contém os arquivos .part1, .part2, etc."
    $labelInstrucoes.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $labelInstrucoes.ForeColor = [System.Drawing.Color]::Gray
    $labelInstrucoes.Size = New-Object System.Drawing.Size(450, 50)
    $labelInstrucoes.Location = New-Object System.Drawing.Point(25, 70)
    $labelInstrucoes.TextAlign = "MiddleCenter"
    $form.Controls.Add($labelInstrucoes)
    
    # Botão Selecionar Pasta
    $btnSelecionar = New-Object System.Windows.Forms.Button
    $btnSelecionar.Text = "SELECIONAR PASTA COM AS PARTES"
    $btnSelecionar.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnSelecionar.ForeColor = [System.Drawing.Color]::White
    $btnSelecionar.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $btnSelecionar.Size = New-Object System.Drawing.Size(400, 60)
    $btnSelecionar.Location = New-Object System.Drawing.Point(50, 130)
    $btnSelecionar.FlatStyle = "Flat"
    $btnSelecionar.FlatAppearance.BorderSize = 0
    $btnSelecionar.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnSelecionar.Add_Click({
        $form.Hide()
        $partsDir = Selecionar-Pasta
        if ($partsDir) {
            Executar-Descompactacao -partsDir $partsDir
        } else {
            $form.Show()
        }
    })
    $form.Controls.Add($btnSelecionar)
    
    # Botão Voltar
    $btnVoltar = New-Object System.Windows.Forms.Button
    $btnVoltar.Text = "VOLTAR AO MENU PRINCIPAL"
    $btnVoltar.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnVoltar.ForeColor = [System.Drawing.Color]::White
    $btnVoltar.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
    $btnVoltar.Size = New-Object System.Drawing.Size(400, 40)
    $btnVoltar.Location = New-Object System.Drawing.Point(50, 210)
    $btnVoltar.FlatStyle = "Flat"
    $btnVoltar.FlatAppearance.BorderSize = 0
    $btnVoltar.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnVoltar.Add_Click({
        $form.Hide()
        Show-MainMenu
        $form.Close()
    })
    $form.Controls.Add($btnVoltar)
    
    $form.ShowDialog() | Out-Null
}

# Função para executar descompactação
function Executar-Descompactacao {
    param([string]$partsDir)
    
    $script:operacaoEmAndamento = $true
    
    try {
        # Verificar arquivos .part*
        $partFiles = @(Get-ChildItem -Path $partsDir -Filter "*.part*" | Sort-Object Name)
        
        if ($partFiles.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Nenhum arquivo .part* encontrado na pasta selecionada!",
                "Aviso",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            Show-DescompactarMenu
            return
        }
        
        # Selecionar pasta de destino
        $dialogDest = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialogDest.Description = "Selecione onde extrair os arquivos"
        $dialogDest.ShowNewFolderButton = $true
        
        if ($dialogDest.ShowDialog() -ne 'OK') {
            Show-DescompactarMenu
            return
        }
        
        $outputFolder = $dialogDest.SelectedPath
        
        # Determinar nome base
        $firstPart = $partFiles[0].Name
        if ($firstPart -match "(.+)\.part\d+$") {
            $baseName = $matches[1]
        } else {
            $baseName = "arquivo_restaurado"
        }
        
        $outputZipPath = Join-Path $outputFolder "$baseName.zip"
        
        # Mostrar tela de progresso
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "Descompactando..."
        $progressForm.Size = New-Object System.Drawing.Size(500, 200)
        $progressForm.StartPosition = "CenterScreen"
        $progressForm.FormBorderStyle = "FixedDialog"
        $progressForm.MaximizeBox = $false
        $progressForm.Topmost = $true
        
        # Label de status
        $labelStatus = New-Object System.Windows.Forms.Label
        $labelStatus.Text = "Juntando partes..."
        $labelStatus.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $labelStatus.Size = New-Object System.Drawing.Size(460, 30)
        $labelStatus.Location = New-Object System.Drawing.Point(20, 20)
        $progressForm.Controls.Add($labelStatus)
        
        # Barra de progresso
        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Size = New-Object System.Drawing.Size(460, 30)
        $progressBar.Location = New-Object System.Drawing.Point(20, 60)
        $progressBar.Style = "Marquee"
        $progressBar.MarqueeAnimationSpeed = 50
        $progressForm.Controls.Add($progressBar)
        
        # Label de detalhes
        $labelDetalhes = New-Object System.Windows.Forms.Label
        $labelDetalhes.Text = "Processando $($partFiles.Count) arquivos..."
        $labelDetalhes.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $labelDetalhes.ForeColor = [System.Drawing.Color]::Gray
        $labelDetalhes.Size = New-Object System.Drawing.Size(460, 50)
        $labelDetalhes.Location = New-Object System.Drawing.Point(20, 100)
        $progressForm.Controls.Add($labelDetalhes)
        
        # Mostrar formulário imediatamente
        $progressForm.Show()
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            # Juntar as partes
            $outputStream = [System.IO.File]::OpenWrite($outputZipPath)
            
            try {
                $count = 0
                foreach ($part in $partFiles) {
                    $count++
                    $labelDetalhes.Text = "Processando $($part.Name) ($count/$($partFiles.Count))"
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    $bytes = [System.IO.File]::ReadAllBytes($part.FullName)
                    $outputStream.Write($bytes, 0, $bytes.Length)
                }
            } finally {
                $outputStream.Close()
            }
            
            # Extrair
            $labelStatus.Text = "Extraindo conteúdo..."
            $labelDetalhes.Text = "Aguarde..."
            [System.Windows.Forms.Application]::DoEvents()
            
            if (-not (Test-Path $outputFolder)) {
                New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
            }
            
            Expand-Archive -Path $outputZipPath -DestinationPath $outputFolder -Force
            
            # Fechar formulário de progresso
            $progressForm.Close()
            
            # Perguntar se quer remover o ZIP
            $resultRemover = [System.Windows.Forms.MessageBox]::Show(
                "Descompactação concluída com sucesso!`n`nDeseja remover o arquivo ZIP temporário?",
                "Concluído",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            
            if ($resultRemover -eq [System.Windows.Forms.DialogResult]::Yes) {
                Remove-Item $outputZipPath -Force -ErrorAction SilentlyContinue
            }
            
            # Abrir pasta de destino
            Start-Process "explorer.exe" -ArgumentList $outputFolder
            
            Show-MainMenu
            
        } catch {
            $progressForm.Close()
            throw $_
        }
        
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Erro durante a descompactação: $($_.Exception.Message)",
            "Erro na Descompactação",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        Show-MainMenu
    } finally {
        $script:operacaoEmAndamento = $false
    }
}

# Função para configurar tamanho das partes
function Show-ConfigurarTamanho {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Configurar Tamanho das Partes"
    $form.Size = New-Object System.Drawing.Size(500, 500)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    
    # Título
    $labelTitulo = New-Object System.Windows.Forms.Label
    $labelTitulo.Text = "CONFIGURAR TAMANHO DAS PARTES"
    $labelTitulo.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $labelTitulo.ForeColor = [System.Drawing.Color]::DarkOrange
    $labelTitulo.Size = New-Object System.Drawing.Size(450, 40)
    $labelTitulo.Location = New-Object System.Drawing.Point(25, 20)
    $labelTitulo.TextAlign = "MiddleCenter"
    $form.Controls.Add($labelTitulo)
    
    # Tamanho atual
    $labelAtual = New-Object System.Windows.Forms.Label
    $labelAtual.Text = "Tamanho atual: $($script:tamanhoParteMB) MB"
    $labelAtual.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $labelAtual.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $labelAtual.Size = New-Object System.Drawing.Size(450, 30)
    $labelAtual.Location = New-Object System.Drawing.Point(25, 70)
    $labelAtual.TextAlign = "MiddleCenter"
    $form.Controls.Add($labelAtual)
    
    # Botão 90 MB (Padrão)
    $btn90 = New-Object System.Windows.Forms.Button
    $btn90.Text = "90 MB (Padrão)"
    $btn90.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btn90.ForeColor = [System.Drawing.Color]::White
    $btn90.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $btn90.Size = New-Object System.Drawing.Size(400, 40)
    $btn90.Location = New-Object System.Drawing.Point(50, 120)
    $btn90.FlatStyle = "Flat"
    $btn90.FlatAppearance.BorderSize = 0
    $btn90.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn90.Add_Click({
        $script:tamanhoParteMB = 90
        $script:tamanhoParteBytes = 90 * 1MB
        $form.Close()
        [System.Windows.Forms.MessageBox]::Show(
            "Tamanho alterado para 90 MB!",
            "Configuração Alterada",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        Show-MainMenu
    })
    $form.Controls.Add($btn90)
    
    # Botão 50 MB
    $btn50 = New-Object System.Windows.Forms.Button
    $btn50.Text = "50 MB"
    $btn50.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $btn50.ForeColor = [System.Drawing.Color]::White
    $btn50.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    $btn50.Size = New-Object System.Drawing.Size(400, 40)
    $btn50.Location = New-Object System.Drawing.Point(50, 170)
    $btn50.FlatStyle = "Flat"
    $btn50.FlatAppearance.BorderSize = 0
    $btn50.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn50.Add_Click({
        $script:tamanhoParteMB = 50
        $script:tamanhoParteBytes = 50 * 1MB
        $form.Close()
        [System.Windows.Forms.MessageBox]::Show(
            "Tamanho alterado para 50 MB!",
            "Configuração Alterada",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        Show-MainMenu
    })
    $form.Controls.Add($btn50)
    
    # Botão 25 MB
    $btn25 = New-Object System.Windows.Forms.Button
    $btn25.Text = "25 MB"
    $btn25.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $btn25.ForeColor = [System.Drawing.Color]::White
    $btn25.BackColor = [System.Drawing.Color]::FromArgb(111, 66, 193)
    $btn25.Size = New-Object System.Drawing.Size(400, 40)
    $btn25.Location = New-Object System.Drawing.Point(50, 220)
    $btn25.FlatStyle = "Flat"
    $btn25.FlatAppearance.BorderSize = 0
    $btn25.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn25.Add_Click({
        $script:tamanhoParteMB = 25
        $script:tamanhoParteBytes = 25 * 1MB
        $form.Close()
        [System.Windows.Forms.MessageBox]::Show(
            "Tamanho alterado para 25 MB!",
            "Configuração Alterada",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        Show-MainMenu
    })
    $form.Controls.Add($btn25)
    
    # Botão 10 MB
    $btn10 = New-Object System.Windows.Forms.Button
    $btn10.Text = "10 MB"
    $btn10.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $btn10.ForeColor = [System.Drawing.Color]::White
    $btn10.BackColor = [System.Drawing.Color]::FromArgb(23, 162, 184)
    $btn10.Size = New-Object System.Drawing.Size(400, 40)
    $btn10.Location = New-Object System.Drawing.Point(50, 270)
    $btn10.FlatStyle = "Flat"
    $btn10.FlatAppearance.BorderSize = 0
    $btn10.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn10.Add_Click({
        $script:tamanhoParteMB = 10
        $script:tamanhoParteBytes = 10 * 1MB
        $form.Close()
        [System.Windows.Forms.MessageBox]::Show(
            "Tamanho alterado para 10 MB!",
            "Configuração Alterada",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        Show-MainMenu
    })
    $form.Controls.Add($btn10)
    
    # Botão Personalizado
    $btnPersonalizado = New-Object System.Windows.Forms.Button
    $btnPersonalizado.Text = "TAMANHO PERSONALIZADO (até 1000 MB)"
    $btnPersonalizado.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $btnPersonalizado.ForeColor = [System.Drawing.Color]::White
    $btnPersonalizado.BackColor = [System.Drawing.Color]::FromArgb(255, 193, 7)
    $btnPersonalizado.Size = New-Object System.Drawing.Size(400, 40)
    $btnPersonalizado.Location = New-Object System.Drawing.Point(50, 320)
    $btnPersonalizado.FlatStyle = "Flat"
    $btnPersonalizado.FlatAppearance.BorderSize = 0
    $btnPersonalizado.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnPersonalizado.Add_Click({
        $inputForm = New-Object System.Windows.Forms.Form
        $inputForm.Text = "Tamanho Personalizado"
        $inputForm.Size = New-Object System.Drawing.Size(350, 220)
        $inputForm.StartPosition = "CenterScreen"
        $inputForm.FormBorderStyle = "FixedDialog"
        $inputForm.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
        
        $label = New-Object System.Windows.Forms.Label
        $label.Text = "Digite o tamanho em MB (1-1000):"
        $label.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $label.Size = New-Object System.Drawing.Size(300, 40)
        $label.Location = New-Object System.Drawing.Point(25, 20)
        $inputForm.Controls.Add($label)
        
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Font = New-Object System.Drawing.Font("Segoe UI", 12)
        $textBox.Size = New-Object System.Drawing.Size(200, 30)
        $textBox.Location = New-Object System.Drawing.Point(70, 70)
        $textBox.Text = $script:tamanhoParteMB
        $textBox.TextAlign = "Center"
        $inputForm.Controls.Add($textBox)
        
        $labelMB = New-Object System.Windows.Forms.Label
        $labelMB.Text = "MB"
        $labelMB.Font = New-Object System.Drawing.Font("Segoe UI", 12)
        $labelMB.Size = New-Object System.Drawing.Size(40, 30)
        $labelMB.Location = New-Object System.Drawing.Point(275, 70)
        $inputForm.Controls.Add($labelMB)
        
        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Text = "OK"
        $btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $btnOK.ForeColor = [System.Drawing.Color]::White
        $btnOK.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
        $btnOK.Size = New-Object System.Drawing.Size(100, 35)
        $btnOK.Location = New-Object System.Drawing.Point(70, 120)
        $btnOK.FlatStyle = "Flat"
        $btnOK.FlatAppearance.BorderSize = 0
        $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $inputForm.AcceptButton = $btnOK
        $inputForm.Controls.Add($btnOK)
        
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = "Cancelar"
        $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $btnCancel.ForeColor = [System.Drawing.Color]::White
        $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
        $btnCancel.Size = New-Object System.Drawing.Size(100, 35)
        $btnCancel.Location = New-Object System.Drawing.Point(180, 120)
        $btnCancel.FlatStyle = "Flat"
        $btnCancel.FlatAppearance.BorderSize = 0
        $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $inputForm.Controls.Add($btnCancel)
        
        if ($inputForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $novoTamanho = $textBox.Text.Trim()
            
            if ([string]::IsNullOrEmpty($novoTamanho)) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Digite um valor para o tamanho!",
                    "Valor Vazio",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return
            }
            
            if (-not [int]::TryParse($novoTamanho, [ref]$null)) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Valor inválido! Digite apenas números.",
                    "Valor Inválido",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return
            }
            
            $tamanhoInt = [int]$novoTamanho
            
            if ($tamanhoInt -le 0 -or $tamanhoInt -gt 1000) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Digite um valor entre 1 e 1000 MB!",
                    "Valor Fora do Intervalo",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                return
            }
            
            $script:tamanhoParteMB = $tamanhoInt
            $script:tamanhoParteBytes = $tamanhoInt * 1MB
            $form.Close()
            [System.Windows.Forms.MessageBox]::Show(
                "Tamanho alterado para $tamanhoInt MB!",
                "Configuração Alterada",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            Show-MainMenu
        }
    })
    $form.Controls.Add($btnPersonalizado)
    
    # Botão Voltar
    $btnVoltar = New-Object System.Windows.Forms.Button
    $btnVoltar.Text = "VOLTAR AO MENU PRINCIPAL"
    $btnVoltar.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnVoltar.ForeColor = [System.Drawing.Color]::White
    $btnVoltar.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
    $btnVoltar.Size = New-Object System.Drawing.Size(400, 40)
    $btnVoltar.Location = New-Object System.Drawing.Point(50, 380)
    $btnVoltar.FlatStyle = "Flat"
    $btnVoltar.FlatAppearance.BorderSize = 0
    $btnVoltar.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnVoltar.Add_Click({
        $form.Hide()
        Show-MainMenu
        $form.Close()
    })
    $form.Controls.Add($btnVoltar)
    
    $form.ShowDialog() | Out-Null
}

# Iniciar aplicação
try {
    # Verificar se PowerShell tem permissões necessárias
    Show-MainMenu
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Erro ao iniciar o aplicativo: $($_.Exception.Message)`n`nTente executar como Administrador.",
        "Erro Fatal",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}