# Diagnóstico de Indisponibilidade — Streamlit (`189.112.77.92:8501`)

**Data:** 10/07/2026  
**Modo:** Auditoria read-only (sem alterações no sistema)  
**Sintoma reportado:** `ERR_CONNECTION_TIMED_OUT` ao acessar `http://189.112.77.92:8501`

---

## Resultado final

| Afirmação | Status |
|-----------|--------|
| A aplicação está saudável | **SIM** |
| O problema é de infraestrutura de rede | **SIM** |
| Nenhuma alteração de código é necessária | **SIM** |
| Próxima etapa: administrador de rede / infraestrutura | **SIM** |

---

## Arquitetura do diagnóstico (3 camadas)

```
[Internet / Cliente externo]
         │
         ▼
┌─────────────────────────────────────────┐
│ CAMADA 3 — Infraestrutura               │
│ NAT, roteador, firewall de borda,       │
│ operadora, hairpin NAT                  │
└─────────────────────────────────────────┘
         │
         ▼
┌─────────────────────────────────────────┐
│ CAMADA 2 — Sistema Operacional (Windows)│
│ Windows Firewall (acesso local à porta) │
└─────────────────────────────────────────┘
         │
         ▼
┌─────────────────────────────────────────┐
│ CAMADA 1 — Aplicação                    │
│ Streamlit + código do projeto           │
└─────────────────────────────────────────┘
```

---

## CAMADA 1 — Aplicação

**Resultado: SAUDÁVEL — sem falha de inicialização**

| Verificação | Evidência | Resultado |
|-------------|-----------|-----------|
| Streamlit em execução | PID 9076 (`streamlit`), PID 16852 (`python`) | ✔ |
| Porta 8501 em LISTENING | `0.0.0.0:8501` e `[::]:8501` — PID 16852 | ✔ |
| Resposta local | `http://localhost:8501` → HTTP **200** | ✔ |
| Resposta na LAN | `http://10.10.0.153:8501` → HTTP **200** | ✔ |
| Código íntegro | `import app` → OK; novos services → OK | ✔ |
| Traceback no terminal | Nenhum | ✔ |
| Erro de import / syntax | Nenhum | ✔ |

**Mensagem de startup (terminal):**
```
You can now view your Streamlit app in your browser.

Local URL:    http://localhost:8501
Network URL:  http://10.10.0.153:8501
External URL: http://189.112.77.92:8501
```

**Configuração Streamlit** (`.streamlit/config.toml`):
- `server.headless = true`
- `server.address` e `server.port` não definidos → padrão Streamlit (bind em todas as interfaces, porta 8501)

**Alterações recentes de código:** não impedem o startup. A aplicação sobe e responde localmente com as alterações de reentrega/complementação presentes no working tree.

**Conclusão da Camada 1:** a indisponibilidade externa **não é causada** por falha da aplicação, crash do Streamlit ou erro de código.

---

## CAMADA 2 — Sistema Operacional (Windows)

**Resultado: VERIFICAR / DOCUMENTAR bloqueios locais**

O Windows controla **apenas o acesso local** à porta via **Windows Firewall**. Port Forwarding **não é função do Windows** neste cenário.

### Firewall Windows — estado geral

| Perfil | Estado |
|--------|--------|
| Domínio | Ligado |
| Particular | Ligado |
| Público | Ligado |

### Regras relevantes documentadas

**Porta 8501 / Streamlit:**
- Nenhuma regra de firewall específica encontrada para porta `8501` ou nome `Streamlit`.

**`python.exe` (Python 3.12 — runtime do Streamlit):**

| Direção | Protocolo | Perfil | Ação | Programa |
|---------|-----------|--------|------|----------|
| Entrada | TCP | Público | **BLOQUEAR** | `python312\python.exe` |
| Entrada | UDP | Público | **BLOQUEAR** | `python312\python.exe` |

**`python.exe` (Python 3.13):**
- Regras de **PERMITIR** entrada TCP/UDP no perfil Público existem, mas o Streamlit em execução utiliza **Python 3.12**.

### Port proxy no Windows

- `netsh interface portproxy show all` → **nenhuma regra configurada**
- Isso é esperado: o Windows não realiza NAT/port forwarding para exposição externa neste cenário.

### Teste de acesso por camada

| Destino | Resultado |
|---------|-----------|
| `localhost:8501` | HTTP 200 ✔ |
| `10.10.0.153:8501` (LAN) | HTTP 200 ✔ |
| `189.112.77.92:8501` (IP público) | **TIMEOUT** ✗ |

**Conclusão da Camada 2:** existe evidência de **bloqueio de entrada TCP para `python.exe` 3.12 no perfil Público**. Isso pode impedir conexões externas que cheguem à máquina. Ação recomendada para o administrador: revisar e, se necessário, criar regra de entrada permitindo TCP 8501 (ou ajustar regra do `python.exe` 3.12).

---

## CAMADA 3 — Infraestrutura

**Resultado: dependência de encaminhamento e políticas externas à aplicação**

O acesso externo via `http://189.112.77.92:8501` depende de componentes **fora do Windows e fora da aplicação**:

| Componente | Papel |
|------------|-------|
| **NAT** | Traduz IP público (`189.112.77.92`) para IP privado (`10.10.0.153`) |
| **Port Forwarding no roteador / firewall de borda** | Encaminha porta externa 8501 → `10.10.0.153:8501` |
| **Firewall corporativo** | Pode bloquear entrada na porta 8501 antes de chegar ao host |
| **Regras da operadora** | CGNAT, bloqueio de portas, políticas de segurança |
| **Hairpin NAT** | Quando o teste é feito da própria rede interna usando o IP público; se ausente, timeout local no IP público é esperado |

### Topologia identificada

```
IP público (Streamlit):  189.112.77.92
IP local (LAN):          10.10.0.153
Porta da aplicação:      8501
```

A máquina está em rede privada (`10.10.0.x`). Clientes na Internet precisam que o tráfego para `189.112.77.92:8501` seja encaminhado pelo **roteador/gateway** até `10.10.0.153:8501`.

> **Correção técnica aplicada neste relatório:**  
> Port Forwarding **não é configuração do Windows** para este cenário.  
> O encaminhamento de portas é realizado em **roteador, firewall de borda, gateway, equipamento NAT ou políticas da operadora**.  
> O Windows apenas controla se a porta é aceita **localmente** após o pacote chegar ao host.

**Conclusão da Camada 3:** o `ERR_CONNECTION_TIMED_OUT` externo é consistente com bloqueio ou ausência de rota na **infraestrutura de rede**, não com falha da aplicação.

---

## Dependências (registro separado)

### Versões em execução

| Componente | Versão observada |
|------------|------------------|
| `streamlit` (CLI em execução) | **1.53.0** (`python -m streamlit --version`) |
| `streamlit` (metadado pip) | 1.51.0 (`pip show streamlit`) |

### Avisos do `pip check`

```
streamlit 1.51.0 has requirement packaging<26,>=20, but you have packaging 26.1
streamlit 1.51.0 has requirement pandas<3,>=1.4.0, but you have pandas 3.0.2
```

### Interpretação

| Pergunta | Resposta |
|----------|----------|
| Esses avisos impediram a inicialização? | **NÃO** — Streamlit iniciou e responde HTTP 200 localmente |
| Relação com timeout externo? | **NENHUMA** — são inconsistências de metadados de dependência, não bloqueio de rede |
| Ação imediata necessária? | Não para resolver o timeout; alinhar versões pode ser feito em manutenção futura |

---

## Respostas consolidadas (checklist original)

| # | Pergunta | Resposta |
|---|----------|----------|
| 1 | O Streamlit iniciou? | **SIM** |
| 2 | A porta 8501 está aberta no host? | **SIM** (LISTENING em `0.0.0.0:8501`) |
| 2b | A porta 8501 é acessível externamente? | **NÃO** (timeout em `189.112.77.92:8501`) |
| 3 | Existe erro no código? | **NÃO** |
| 4 | Existe problema de rede? | **SIM** |
| 5 | Existe problema de configuração? | **SIM** — na camada de rede/firewall (Camadas 2 e 3), não no código |
| 6 | Causa raiz | Bloqueio/ausência de rota na infraestrutura de rede entre Internet e porta 8501 do host |
| 7 | Correção recomendada | Conduzida pelo administrador de rede (ver próxima seção) |

---

## Próxima etapa — Administrador de rede / infraestrutura

**Nenhuma alteração de código é necessária.**

Checklist para o responsável pela infraestrutura:

- [ ] **Firewall Windows** — verificar regra de bloqueio de entrada TCP para `python.exe` 3.12 (perfil Público); criar regra de permissão para porta 8501 se necessário
- [ ] **Firewall corporativo** — confirmar se porta 8501 está liberada para entrada
- [ ] **NAT / Roteador / Gateway** — configurar Port Forwarding: `8501` externo → `10.10.0.153:8501`
- [ ] **Políticas da operadora** — verificar CGNAT ou bloqueio de portas de entrada
- [ ] **Hairpin NAT** — se testes forem feitos de dentro da mesma rede usando IP público, validar suporte ou testar de rede externa real
- [ ] **Alternativa segura** — considerar VPN, túnel (Cloudflare Tunnel, ngrok) ou reverse proxy com HTTPS em vez de exposição direta da porta 8501

---

## Evidências coletadas (referência)

| Comando / verificação | Resultado |
|-----------------------|-----------|
| `netstat -ano \| findstr :8501` | LISTENING PID 16852 em `0.0.0.0:8501` |
| `Invoke-WebRequest localhost:8501` | HTTP 200 |
| `Invoke-WebRequest 10.10.0.153:8501` | HTTP 200 |
| `Invoke-WebRequest 189.112.77.92:8501` | Timeout |
| `ipconfig` | IPv4: `10.10.0.153` |
| `import app` | OK |
| `pip check` | Avisos packaging/pandas (sem impacto no startup) |

---

*Relatório revisado conforme solicitação de correção técnica sobre Port Forwarding e separação em camadas (Aplicação / SO / Infraestrutura).*
