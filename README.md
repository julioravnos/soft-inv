# Windows Server Software Inventory / Inventário de Software em Servidores Windows

**🌐 Language / Idioma:** [English](#-english) · [Português (BR)](#-português-br)

---

# 🇺🇸 English

A single-file PowerShell tool with a graphical interface that collects **installed applications and versions** from **remote Windows servers (Windows Server 2008 and later)** and exports the results to **CSV, HTML or PDF**.

No agents, no installation on the servers, no external modules.

## ✨ Features

- **Simple GUI** – type or paste the server list (one per line), or import it from a `.txt` file.
- **Parallel collection** – query many servers at the same time (configurable, 1–50 simultaneous servers) with a progress bar and a cancel button.
- **Windows updates filter** – hide Windows updates (KB, hotfixes, security/cumulative updates) with a single checkbox, toggleable at any time without re-collecting.
- **Three export formats** – CSV, HTML (with a built-in search box) and PDF. Exports respect the filter currently active on screen.
- **Agentless** – reads the remote registry through WMI (`StdRegProv`). It does **not** depend on the *Remote Registry* service.
- **Fast and safe** – avoids `Win32_Product`, which is slow and can trigger MSI reconfiguration on the servers.
- **Resilient** – offline or unreachable servers are reported as `[FALHA]` rows with the reason, without stopping the rest of the run.
- **32-bit and 64-bit applications** are both collected.
- **Alternate credentials** – optionally run the collection with a different administrative account.

## 📋 Collected data

| Column | Description |
|---|---|
| Servidor | Server name or IP |
| Aplicativo | Application name |
| Versao | Application version |
| Fabricante | Publisher |
| DataInstalacao | Install date (when registered by the installer) |
| Arquitetura | 32 bits or 64 bits |
| DataColeta | Date/time of the collection |

## ✅ Requirements

- **Workstation running the script:** Windows with PowerShell 3.0+ (5.1 recommended).
- **Target servers:** Windows Server 2008 or later.
- A user with **administrator permissions** on the target servers.
- **TCP port 135 (RPC/DCOM)** and the dynamic WMI ports reachable from the workstation (default WMI requirements).
- **PDF export only:** Microsoft Edge or Google Chrome installed on the workstation (used in headless mode).

## 🚀 Usage

1. Download `Inventario-Software.ps1`.
2. Right-click the file → **Run with PowerShell**.
   If execution is blocked by policy:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Inventario-Software.ps1
   ```
3. Enter the servers (name or IP, one per line) or click **Importar lista (.txt)**. Lines starting with `#` are ignored.
4. *(Optional)* Check **Usar credenciais alternativas** if your current user is not an administrator on the servers.
5. Adjust **Servidores em paralelo** if needed and click **Coletar**.
6. Use **Exportar CSV**, **Exportar HTML** or **Exportar PDF**.

### Example server list

```text
# Database servers
srv-dados01
srv-dados02
192.168.10.25
```

## ⚙️ Configuration

At the top of the script:

```powershell
$Delimitador       = ';'   # CSV delimiter. ';' opens directly in Excel pt-BR; use ',' otherwise
$TimeoutPortaMs    = 3000  # Timeout (ms) for the port 135 check on each server
$ParalelismoPadrao = 10    # Default number of simultaneous servers
```

## 🧠 How it works

1. For each server, the script checks that port 135 is reachable (short timeout, so offline servers do not stall the run).
2. It connects to the WMI `StdRegProv` class in `root\default`.
3. It reads the `Uninstall` registry keys (64-bit and `Wow6432Node`) under `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion`.
4. Each server runs in its own runspace, so many servers are processed in parallel while the GUI stays responsive.

## 🛠️ Troubleshooting

| Message | Likely cause |
|---|---|
| `Porta 135 inacessivel ou servidor offline` | Server is off, name does not resolve, or a firewall blocks RPC |
| `Acesso negado` / `Access denied` | The user is not an administrator on the server – try alternate credentials |
| `O servidor RPC não está disponível` | RPC/WMI blocked by firewall or WMI service not running |
| PDF export fails | Edge or Chrome not found on the workstation – export to HTML and print to PDF instead |

## ⚠️ Notes

- Very high parallelism can create many simultaneous RPC connections; start with 10–15 and increase gradually.
- The KB filter is name-based (e.g. `KB1234567`, "Security Update", "Hotfix", "Update Rollup", "Cumulative Update", "Update for Windows/Microsoft"). Some updates may not match; adjust `Test-Atualizacao` in the script if needed.
- Only software registered in the Windows `Uninstall` keys is listed (portable applications are not detected).
- Use only on servers you are authorized to inventory.


---

# 🇧🇷 Português (BR)

Ferramenta em PowerShell, em arquivo único e com interface gráfica, que coleta **aplicativos e versões instalados** em **servidores Windows remotos (Windows Server 2008 ou superior)** e exporta o resultado em **CSV, HTML ou PDF**.

Sem agentes, sem instalação nos servidores e sem módulos externos.

## ✨ Funcionalidades

- **Interface gráfica simples** – digite ou cole a lista de servidores (um por linha) ou importe de um arquivo `.txt`.
- **Coleta em paralelo** – consulta vários servidores ao mesmo tempo (configurável, de 1 a 50 simultâneos), com barra de progresso e botão de cancelar.
- **Filtro de atualizações do Windows** – oculta atualizações (KB, hotfixes, atualizações de segurança/cumulativas) com um único checkbox, que pode ser alternado a qualquer momento sem refazer a coleta.
- **Três formatos de exportação** – CSV, HTML (com campo de busca embutido) e PDF. A exportação respeita o filtro ativo na tela.
- **Sem agente** – lê o registro remoto via WMI (`StdRegProv`). **Não** depende do serviço *Remote Registry*.
- **Rápido e seguro** – evita o `Win32_Product`, que é lento e pode disparar a reconfiguração de instaladores MSI nos servidores.
- **Resiliente** – servidores offline ou inacessíveis aparecem como linhas `[FALHA]` com o motivo, sem interromper o restante da coleta.
- **Aplicativos de 32 e 64 bits** são coletados.
- **Credenciais alternativas** – opção de executar a coleta com outra conta administrativa.

## 📋 Dados coletados

| Coluna | Descrição |
|---|---|
| Servidor | Nome ou IP do servidor |
| Aplicativo | Nome do aplicativo |
| Versao | Versão do aplicativo |
| Fabricante | Fabricante (Publisher) |
| DataInstalacao | Data de instalação (quando registrada pelo instalador) |
| Arquitetura | 32 bits ou 64 bits |
| DataColeta | Data/hora da coleta |

## ✅ Requisitos

- **Estação que executa o script:** Windows com PowerShell 3.0 ou superior (recomendado 5.1).
- **Servidores de destino:** Windows Server 2008 ou superior.
- Usuário com **permissão de administrador** nos servidores de destino.
- **Porta TCP 135 (RPC/DCOM)** e portas dinâmicas do WMI acessíveis a partir da estação (requisitos padrão do WMI).
- **Somente para exportar PDF:** Microsoft Edge ou Google Chrome instalado na estação (usado em modo headless).

## 🚀 Como usar

1. Baixe o arquivo `Inventario-Software.ps1`.
2. Clique com o botão direito → **Executar com o PowerShell**.
   Se a execução for bloqueada pela política do sistema:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Inventario-Software.ps1
   ```
3. Informe os servidores (nome ou IP, um por linha) ou clique em **Importar lista (.txt)**. Linhas iniciadas com `#` são ignoradas.
4. *(Opcional)* Marque **Usar credenciais alternativas** se o seu usuário atual não for administrador nos servidores.
5. Ajuste **Servidores em paralelo**, se necessário, e clique em **Coletar**.
6. Use **Exportar CSV**, **Exportar HTML** ou **Exportar PDF**.

### Exemplo de lista de servidores

```text
# Servidores de banco de dados
srv-dados01
srv-dados02
192.168.10.25
```

## ⚙️ Configuração

No início do script:

```powershell
$Delimitador       = ';'   # Delimitador do CSV. ';' abre direto no Excel pt-BR; use ',' caso contrário
$TimeoutPortaMs    = 3000  # Tempo máximo (ms) do teste da porta 135 em cada servidor
$ParalelismoPadrao = 10    # Quantidade padrão de servidores simultâneos
```

## 🧠 Como funciona

1. Para cada servidor, o script testa se a porta 135 está acessível (timeout curto, para que servidores offline não travem a coleta).
2. Conecta à classe WMI `StdRegProv` em `root\default`.
3. Lê as chaves `Uninstall` do registro (64 bits e `Wow6432Node`) em `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion`.
4. Cada servidor roda em um runspace próprio, então vários são processados em paralelo enquanto a interface continua responsiva.

## 🛠️ Solução de problemas

| Mensagem | Causa provável |
|---|---|
| `Porta 135 inacessivel ou servidor offline` | Servidor desligado, nome não resolve ou firewall bloqueando o RPC |
| `Acesso negado` | O usuário não é administrador no servidor – tente credenciais alternativas |
| `O servidor RPC não está disponível` | RPC/WMI bloqueado por firewall ou serviço WMI parado |
| Falha ao exportar PDF | Edge ou Chrome não encontrado na estação – exporte em HTML e imprima como PDF |

## ⚠️ Observações

- Um paralelismo muito alto pode gerar muitas conexões RPC simultâneas; comece com 10 a 15 e aumente aos poucos.
- O filtro de KB é baseado no nome (ex.: `KB1234567`, "Security Update", "Hotfix", "Update Rollup", "Cumulative Update", "Update for Windows/Microsoft"). Algumas atualizações podem não ser reconhecidas; ajuste a função `Test-Atualizacao` no script se necessário.
- Somente softwares registrados nas chaves `Uninstall` do Windows são listados (aplicativos portáteis não são detectados).
- Use apenas em servidores que você tem autorização para inventariar.
