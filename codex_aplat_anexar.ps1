param(
  [Parameter(Mandatory = $true)]
  [int]$Port,

  [Parameter(Mandatory = $true)]
  [string]$AplatHost,

  [Parameter(Mandatory = $true)]
  [string]$EmissionUrl,

  [Parameter(Mandatory = $true)]
  [string]$NumeroPt,

  [Parameter(Mandatory = $true)]
  [string]$PdfPath,

  [Parameter(Mandatory = $true)]
  [string]$LogPath
)

$ErrorActionPreference = 'Stop'

$XPathExibirOpcoes = '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/div[2]/form/app-expando-filter/button'
$XPathCampoNumeroPt = '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/div[2]/form/app-expando-filter/div/div[11]/div[1]/div/div/input'
$XPathBotaoPesquisar = '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/div[2]/form/div/div[5]/div/button'
$XPathCardPt = '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/div[3]/div'
$XPathCardPtDetalhe1 = '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/div[3]/div/ul/li[2]/div/div[1]/div[2]/div[2]'
$XPathCardPtDetalhe2 = '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/div[3]/div/ul/li[2]/div/div[1]/div[2]'
$XPathCardPtPorId1 = "//*[@id='pt_0']/div/div[1]/div[2]/div[2]"
$XPathCardPtPorId2 = "//*[@id='pt_0']/div/div[1]/div[2]"
$XPathCardPtPorId3 = "//*[@id='pt_0']/div"
$XPathCardPtPorId4 = "//*[@id='pt_0']"
$XPathAbaAnexos = '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/app-dynamic-component/app-form-cadastrar-pt/app-modal/div/div/div/div[2]/div/app-cadastrar-pt/app-tabs/ul/li[13]/a'
$XPathBotaoAdicionarAnexo = '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/app-dynamic-component/app-form-cadastrar-pt/app-modal/div/div/div/div[2]/div/app-cadastrar-pt/app-tabs/app-tab[13]/div/app-anexos/div/div/div/app-fileuploader/label/input'
$XPathBotaoAdicionarAnexoAlt1 = "(//app-modal//app-anexos//input[@type='file'])[last()]"
$XPathBotaoAdicionarAnexoAlt2 = "(//app-anexos//input[contains(@class,'upload') or @type='file'])[last()]"
$TextoAvisoArquivoNaoInformado = 'O arquivo para upload não informado.'
$XPathMessageBoxConfirmacao = '/html/body/app-root/div/app-messagebox'
$XPathBotaoOkConfirmacao = '/html/body/app-root/div/app-messagebox/div/div/div/div[3]/div/button'
$XPathBotaoOkConfirmacaoAlt = "(//app-messagebox//button[normalize-space()='Ok' or normalize-space()='OK'])[last()]"
$XPathFecharModal = '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/app-dynamic-component/app-form-cadastrar-pt/app-modal/div/div/div/div[1]/button'

$script:LogBuffer = New-Object System.Collections.Generic.List[string]
$script:MessageId = 0

function Write-Log {
  param([string]$Message)
  $line = ('{0} - {1}' -f (Get-Date).ToString('HH:mm:ss'), $Message)
  $script:LogBuffer.Add($line) | Out-Null
}

function Save-Log {
  $dir = Split-Path -Parent $LogPath
  if ($dir -and -not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
  }
  Set-Content -LiteralPath $LogPath -Value $script:LogBuffer -Encoding UTF8
}

function Escape-Js {
  param([string]$Value)
  $escaped = $Value.Replace('\', '\\')
  $escaped = $escaped.Replace("'", '\x27')
  $escaped = $escaped.Replace("`r", '\r')
  $escaped = $escaped.Replace("`n", '\n')
  return $escaped
}

function Receive-CdpMessage {
  param(
    [System.Net.WebSockets.ClientWebSocket]$Socket,
    [int]$TimeoutMs = 30000
  )

  $buffer = New-Object byte[] 32768
  $segment = [ArraySegment[byte]]::new($buffer)
  $stream = New-Object System.IO.MemoryStream
  $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMs)

  while ($true) {
    $remaining = [int][Math]::Max(1000, ($deadline - [DateTime]::UtcNow).TotalMilliseconds)
    if ([DateTime]::UtcNow -ge $deadline) {
      throw 'Timeout aguardando mensagem do DevTools Protocol.'
    }

    $cts = [System.Threading.CancellationTokenSource]::new($remaining)
    try {
      $result = $Socket.ReceiveAsync($segment, $cts.Token).GetAwaiter().GetResult()
    }
    finally {
      $cts.Dispose()
    }

    if ($result.MessageType -eq [System.Net.WebSockets.WebSocketMessageType]::Close) {
      throw 'A conexão com o DevTools Protocol foi encerrada.'
    }

    if ($result.Count -gt 0) {
      $stream.Write($buffer, 0, $result.Count)
    }

    if ($result.EndOfMessage) {
      break
    }
  }

  $json = [System.Text.Encoding]::UTF8.GetString($stream.ToArray())
  $stream.Dispose()
  if ([string]::IsNullOrWhiteSpace($json)) {
    return $null
  }
  return ($json | ConvertFrom-Json)
}

function Invoke-Cdp {
  param(
    [System.Net.WebSockets.ClientWebSocket]$Socket,
    [string]$Method,
    [hashtable]$Params = @{},
    [int]$TimeoutMs = 30000
  )

  $script:MessageId++
  $payload = @{
    id     = $script:MessageId
    method = $Method
    params = $Params
  } | ConvertTo-Json -Compress -Depth 20

  $bytes = [System.Text.Encoding]::UTF8.GetBytes($payload)
  $segment = [ArraySegment[byte]]::new($bytes)
  $Socket.SendAsync(
    $segment,
    [System.Net.WebSockets.WebSocketMessageType]::Text,
    $true,
    [System.Threading.CancellationToken]::None
  ).GetAwaiter().GetResult()

  $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMs)
  while ([DateTime]::UtcNow -lt $deadline) {
    $remaining = [int][Math]::Max(1000, ($deadline - [DateTime]::UtcNow).TotalMilliseconds)
    $message = Receive-CdpMessage -Socket $Socket -TimeoutMs $remaining
    if ($null -eq $message) {
      continue
    }
    if ($message.PSObject.Properties.Name -contains 'id') {
      if ([int]$message.id -eq $script:MessageId) {
        if ($message.PSObject.Properties.Name -contains 'error') {
          throw ('CDP {0} falhou: {1}' -f $Method, ($message.error | ConvertTo-Json -Compress -Depth 10))
        }
        return $message.result
      }
    }
  }

  throw ('Timeout aguardando resposta do comando CDP {0}.' -f $Method)
}

function Invoke-Evaluate {
  param(
    [System.Net.WebSockets.ClientWebSocket]$Socket,
    [string]$Expression,
    [bool]$ReturnByValue = $true,
    [int]$TimeoutMs = 30000
  )

  return Invoke-Cdp -Socket $Socket -Method 'Runtime.evaluate' -TimeoutMs $TimeoutMs -Params @{
    expression    = $Expression
    awaitPromise  = $true
    returnByValue = $ReturnByValue
    userGesture   = $true
  }
}

function Get-TargetTab {
  $lastError = ''
  for ($attempt = 1; $attempt -le 20; $attempt++) {
    try {
      $tabs = Invoke-RestMethod -Uri ("http://127.0.0.1:{0}/json/list" -f $Port) -UseBasicParsing
      $targets = @($tabs | Where-Object { $_.type -eq 'page' })
      if (-not $targets -or $targets.Count -eq 0) {
        throw 'Nenhuma aba do Edge exposta pelo DevTools Protocol.'
      }

      $target = $targets |
        Where-Object {
          $_.url -like "*$AplatHost*" -or
          $_.title -like '*APLAT*' -or
          $_.title -like '*PTs da Instalação SMAR*'
        } |
        Select-Object -First 1

      if ($null -eq $target) {
        $target = $targets | Select-Object -First 1
      }

      return $target
    }
    catch {
      $lastError = $_.Exception.Message
      if ($attempt -eq 1) {
        Write-Log ('Aguardando o DevTools Protocol responder na porta {0}.' -f $Port)
      }
      Start-Sleep -Milliseconds 1000
    }
  }

  throw ('Nao consegui acessar o Edge na porta {0}. Ultimo erro: {1}' -f $Port, $lastError)
}

function Invoke-JsClickXPath {
  param(
    [System.Net.WebSockets.ClientWebSocket]$Socket,
    [string]$XPath,
    [bool]$DoubleClick = $false,
    [int]$TimeoutMs = 20000
  )

  $dbl = if ($DoubleClick) { 'true' } else { 'false' }
  $js = @"
(async () => {
  const xp = '$(Escape-Js $XPath)';
  const dbl = $dbl;
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const find = () => document.evaluate(xp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
  let el = null;
  for (let i = 0; i < 80; i++) {
    el = find();
    if (el) break;
    await sleep(250);
  }
  if (!el) throw new Error('XPath nao encontrado: ' + xp);
  el.scrollIntoView({ block: 'center', inline: 'center' });
  try { el.focus(); } catch (e) {}
  const rect = el.getBoundingClientRect();
  const mouseOptions = {
    bubbles: true,
    cancelable: true,
    composed: true,
    view: window,
    button: 0,
    buttons: 1,
    clientX: Math.floor(rect.left + (rect.width / 2)),
    clientY: Math.floor(rect.top + (rect.height / 2))
  };
  const firePointer = type => {
    if (window.PointerEvent) {
      el.dispatchEvent(new PointerEvent(type, mouseOptions));
    }
  };
  const fireMouse = (type, detail) => {
    el.dispatchEvent(new MouseEvent(type, Object.assign({ detail: detail }, mouseOptions)));
  };
  firePointer('pointerover');
  fireMouse('mouseover', 0);
  firePointer('pointerdown');
  fireMouse('mousedown', 1);
  firePointer('pointerup');
  fireMouse('mouseup', 1);
  fireMouse('click', 1);
  try { el.click(); } catch (e) {}
  if (dbl) {
    firePointer('pointerdown');
    fireMouse('mousedown', 2);
    firePointer('pointerup');
    fireMouse('mouseup', 2);
    fireMouse('click', 2);
    fireMouse('dblclick', 2);
  }
  return true;
})()
"@
  [void](Invoke-Evaluate -Socket $Socket -Expression $js -ReturnByValue $true -TimeoutMs $TimeoutMs)
}

function Test-XPathExists {
  param(
    [System.Net.WebSockets.ClientWebSocket]$Socket,
    [string]$XPath,
    [int]$TimeoutMs = 5000
  )

  $js = @"
(async () => {
  const xp = '$(Escape-Js $XPath)';
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const find = () => document.evaluate(xp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
  const loops = Math.max(1, Math.ceil($TimeoutMs / 250));
  for (let i = 0; i < loops; i++) {
    if (find()) return true;
    await sleep(250);
  }
  return false;
})()
"@
  $result = Invoke-Evaluate -Socket $Socket -Expression $js -ReturnByValue $true -TimeoutMs ($TimeoutMs + 3000)
  return [bool]$result.result.value
}

function Test-VisibleTextContains {
  param(
    [System.Net.WebSockets.ClientWebSocket]$Socket,
    [string]$Text,
    [int]$TimeoutMs = 2000
  )

  $js = @"
(async () => {
  const expected = '$(Escape-Js $Text)';
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const loops = Math.max(1, Math.ceil($TimeoutMs / 250));
  for (let i = 0; i < loops; i++) {
    const text = (document.body && document.body.innerText) ? document.body.innerText : '';
    if (text.indexOf(expected) >= 0) return true;
    await sleep(250);
  }
  return false;
})()
"@
  $result = Invoke-Evaluate -Socket $Socket -Expression $js -ReturnByValue $true -TimeoutMs ($TimeoutMs + 3000)
  return [bool]$result.result.value
}

function Dismiss-UploadWarningIfPresent {
  param(
    [System.Net.WebSockets.ClientWebSocket]$Socket
  )

  $warningVisible = Test-VisibleTextContains -Socket $Socket -Text $TextoAvisoArquivoNaoInformado -TimeoutMs 1200
  $messageBoxVisible = Test-XPathExists -Socket $Socket -XPath $XPathMessageBoxConfirmacao -TimeoutMs 1200
  if ((-not $warningVisible) -and (-not $messageBoxVisible)) {
    return $false
  }

  if ($warningVisible) {
    Write-Log ('Aviso exibido pelo APLAT: {0}' -f $TextoAvisoArquivoNaoInformado)
  }
  else {
    Write-Log 'Fechando messagebox de confirmação pendente do APLAT.'
  }

  if (Test-XPathExists -Socket $Socket -XPath $XPathBotaoOkConfirmacao -TimeoutMs 2500) {
    Invoke-JsClickXPath -Socket $Socket -XPath $XPathBotaoOkConfirmacao -TimeoutMs 6000
    Start-Sleep -Milliseconds 800
  }
  elseif (Test-XPathExists -Socket $Socket -XPath $XPathBotaoOkConfirmacaoAlt -TimeoutMs 2500) {
    Invoke-JsClickXPath -Socket $Socket -XPath $XPathBotaoOkConfirmacaoAlt -TimeoutMs 6000
    Start-Sleep -Milliseconds 800
  }
  return $true
}

function Open-PtResult {
  param(
    [System.Net.WebSockets.ClientWebSocket]$Socket
  )

  $targets = @(
    $XPathCardPtDetalhe1,
    $XPathCardPtDetalhe2,
    $XPathCardPtPorId1,
    $XPathCardPtPorId2,
    $XPathCardPtPorId3,
    $XPathCardPtPorId4,
    $XPathCardPt
  )

  foreach ($targetXPath in $targets) {
    try {
      Write-Log ('Tentando abrir a PT com duplo clique em: {0}' -f $targetXPath)
      Invoke-JsClickXPath -Socket $Socket -XPath $targetXPath -DoubleClick $true -TimeoutMs 8000
      Start-Sleep -Milliseconds 1800
      if (Test-XPathExists -Socket $Socket -XPath $XPathAbaAnexos -TimeoutMs 5000) {
        Write-Log 'Detalhe da PT aberto.'
        return
      }
    }
    catch {
      Write-Log ('Falha neste alvo de duplo clique: ' + $_.Exception.Message)
    }
  }

  throw 'Nao consegui abrir a PT com duplo clique no card do resultado.'
}

function Invoke-JsSetInputXPath {
  param(
    [System.Net.WebSockets.ClientWebSocket]$Socket,
    [string]$XPath,
    [string]$Value,
    [int]$TimeoutMs = 20000
  )

  $js = @"
(async () => {
  const xp = '$(Escape-Js $XPath)';
  const value = '$(Escape-Js $Value)';
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const find = () => document.evaluate(xp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
  let el = null;
  for (let i = 0; i < 80; i++) {
    el = find();
    if (el) break;
    await sleep(250);
  }
  if (!el) throw new Error('Campo nao encontrado: ' + xp);
  el.scrollIntoView({ block: 'center', inline: 'center' });
  try { el.focus(); } catch (e) {}
  const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;
  if (setter) setter.call(el, '');
  else el.value = '';
  el.dispatchEvent(new Event('input', { bubbles: true }));
  if (setter) setter.call(el, value);
  else el.value = value;
  el.dispatchEvent(new Event('input', { bubbles: true }));
  el.dispatchEvent(new Event('change', { bubbles: true }));
  return true;
})()
"@
  [void](Invoke-Evaluate -Socket $Socket -Expression $js -ReturnByValue $true -TimeoutMs $TimeoutMs)
}

function Set-FileInputByXPath {
  param(
    [System.Net.WebSockets.ClientWebSocket]$Socket,
    [string]$XPath,
    [string]$FilePath,
    [int]$TimeoutMs = 20000
  )

  $js = @"
(async () => {
  const xp = '$(Escape-Js $XPath)';
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const find = () => document.evaluate(xp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
  let el = null;
  for (let i = 0; i < 80; i++) {
    el = find();
    if (el) break;
    await sleep(250);
  }
  if (!el) throw new Error('Input de arquivo nao encontrado: ' + xp);
  el.scrollIntoView({ block: 'center', inline: 'center' });
  try { el.focus(); } catch (e) {}
  return el;
})()
"@

  $result = Invoke-Evaluate -Socket $Socket -Expression $js -ReturnByValue $false -TimeoutMs $TimeoutMs
  $objectId = $result.result.objectId
  if ([string]::IsNullOrWhiteSpace($objectId)) {
    throw 'Nao foi possivel obter objectId do input file.'
  }

  [void](Invoke-Cdp -Socket $Socket -Method 'DOM.setFileInputFiles' -TimeoutMs $TimeoutMs -Params @{
    files    = @($FilePath)
    objectId = $objectId
  })

  $jsChange = @"
(async () => {
  const xp = '$(Escape-Js $XPath)';
  const el = document.evaluate(xp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
  if (!el) throw new Error('Input de arquivo sumiu antes do evento change.');
  el.dispatchEvent(new Event('input', { bubbles: true }));
  el.dispatchEvent(new Event('change', { bubbles: true }));
  const fileCount = el.files ? el.files.length : 0;
  const fileName = (fileCount > 0 && el.files[0]) ? el.files[0].name : '';
  return { fileCount, fileName };
})()
"@
  return Invoke-Evaluate -Socket $Socket -Expression $jsChange -ReturnByValue $true -TimeoutMs $TimeoutMs
}

function Wait-ForAttachmentResult {
  param(
    [System.Net.WebSockets.ClientWebSocket]$Socket,
    [string]$ExpectedFileName,
    [int]$TimeoutMs = 25000
  )

  $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMs)
  $warningSeen = $false

  while ([DateTime]::UtcNow -lt $deadline) {
    if (Dismiss-UploadWarningIfPresent -Socket $Socket) {
      $warningSeen = $true
    }

    if (Test-VisibleTextContains -Socket $Socket -Text $ExpectedFileName -TimeoutMs 1000) {
      Write-Log ('Anexo confirmado na tela: {0}' -f $ExpectedFileName)
      [void](Dismiss-UploadWarningIfPresent -Socket $Socket)
      return $true
    }

    Start-Sleep -Milliseconds 700
  }

  if ($warningSeen) {
    Write-Log 'O aviso de arquivo não informado apareceu e o anexo não foi confirmado na tela.'
  }

  return $false
}

function Set-FileInputWithFallback {
  param(
    [System.Net.WebSockets.ClientWebSocket]$Socket,
    [string]$FilePath,
    [int]$TimeoutMs = 25000
  )

  $fileName = Split-Path -Leaf $FilePath
  $targets = @(
    $XPathBotaoAdicionarAnexo,
    $XPathBotaoAdicionarAnexoAlt1,
    $XPathBotaoAdicionarAnexoAlt2
  ) | Select-Object -Unique

  foreach ($targetXPath in $targets) {
    try {
      Write-Log ('Tentando definir anexo pelo input: {0}' -f $targetXPath)
      $state = Set-FileInputByXPath -Socket $Socket -XPath $targetXPath -FilePath $FilePath -TimeoutMs $TimeoutMs
      $fileCount = 0
      $selectedName = ''
      if ($state -and $state.result -and $state.result.value) {
        $fileCount = [int]$state.result.value.fileCount
        $selectedName = [string]$state.result.value.fileName
      }
      Write-Log ('Input recebeu {0} arquivo(s). Nome detectado: {1}' -f $fileCount, $selectedName)

      if (Wait-ForAttachmentResult -Socket $Socket -ExpectedFileName $fileName -TimeoutMs 20000) {
        return
      }
    }
    catch {
      Write-Log ('Falha ao informar arquivo neste input: ' + $_.Exception.Message)
    }
  }

  throw ('Nao consegui confirmar o anexo do arquivo {0} na tela.' -f $fileName)
}

try {
  if (-not (Test-Path -LiteralPath $PdfPath)) {
    throw ('Arquivo PDF nao encontrado: {0}' -f $PdfPath)
  }

  Write-Log ('Conectando ao Edge na porta {0}.' -f $Port)
  $target = Get-TargetTab
  if ($null -eq $target -or [string]::IsNullOrWhiteSpace($target.webSocketDebuggerUrl)) {
    throw 'Nao encontrei uma aba do APLAT exposta pelo DevTools Protocol.'
  }

  $socket = [System.Net.WebSockets.ClientWebSocket]::new()
  $socket.ConnectAsync([Uri]$target.webSocketDebuggerUrl, [System.Threading.CancellationToken]::None).GetAwaiter().GetResult()
  try {
    [void](Invoke-Cdp -Socket $socket -Method 'Page.enable')
    [void](Invoke-Cdp -Socket $socket -Method 'Runtime.enable')
    [void](Invoke-Cdp -Socket $socket -Method 'DOM.enable')

    Write-Log 'Navegando para Emissao de PT.'
    [void](Invoke-Cdp -Socket $socket -Method 'Page.navigate' -Params @{ url = $EmissionUrl } -TimeoutMs 30000)
    Start-Sleep -Milliseconds 4500

    Write-Log 'Abrindo filtros.'
    Invoke-JsClickXPath -Socket $socket -XPath $XPathExibirOpcoes -TimeoutMs 20000
    Start-Sleep -Milliseconds 1200

    Write-Log ('Preenchendo PT {0}.' -f $NumeroPt)
    Invoke-JsSetInputXPath -Socket $socket -XPath $XPathCampoNumeroPt -Value $NumeroPt -TimeoutMs 20000
    Start-Sleep -Milliseconds 700

    Write-Log 'Pesquisando PT.'
    Invoke-JsClickXPath -Socket $socket -XPath $XPathBotaoPesquisar -TimeoutMs 20000
    Start-Sleep -Milliseconds 3000

    Write-Log 'Abrindo a PT.'
    Open-PtResult -Socket $socket
    Start-Sleep -Milliseconds 1500

    Write-Log 'Abrindo aba Anexos.'
    Invoke-JsClickXPath -Socket $socket -XPath $XPathAbaAnexos -TimeoutMs 20000
    Start-Sleep -Milliseconds 1800

    Write-Log 'Definindo arquivo do anexo.'
    Set-FileInputWithFallback -Socket $socket -FilePath $PdfPath -TimeoutMs 25000
    Start-Sleep -Milliseconds 1500
    [void](Dismiss-UploadWarningIfPresent -Socket $socket)

    Write-Log 'Fechando modal da PT.'
    Invoke-JsClickXPath -Socket $socket -XPath $XPathFecharModal -TimeoutMs 15000
    Start-Sleep -Milliseconds 1200
    [void](Dismiss-UploadWarningIfPresent -Socket $socket)

    Write-Log 'Fluxo concluido com sucesso.'
  }
  finally {
    if ($socket) {
      if ($socket.State -eq [System.Net.WebSockets.WebSocketState]::Open) {
        $socket.CloseAsync(
          [System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure,
          'ok',
          [System.Threading.CancellationToken]::None
        ).GetAwaiter().GetResult()
      }
      $socket.Dispose()
    }
  }

  Save-Log
  exit 0
}
catch {
  Write-Log ('ERRO: ' + $_.Exception.Message)
  Save-Log
  Write-Error $_
  exit 1
}
