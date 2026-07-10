# Relatório Técnico — Correção Arquitetural do Fluxo de Reentrega e Complementação

**Data:** 10/07/2026  
**Método:** GO • RCA • Arquitetura Senior  
**Status:** Implementado e homologado por testes automatizados

---

## 1. RCA — Causa Raiz

### Incidente
Durante complementação de minuta, lote misto (NFs novas + NF com histórico) gerou:

```
IntegrityError: UNIQUE constraint failed: item_carregamento
(carregamento_id, numero_nf, codigo_produto, sequencia)
```

### Cadeia causal
1. `FechamentoCarregamentoService._complementar_carregamento()` filtrava NFs por classificação de lote (`ClassificacaoNfLote.NOVA`).
2. NFs existentes no mesmo carregamento podiam entrar no fluxo de INSERT por falha de classificação ou deduplicação incompleta.
3. A deduplicação usava chave `(nf, chave_nfe, cprod)` sem validação alinhada à UNIQUE real do banco.
4. `except Exception` convertia `IntegrityError` em falha monolítica do lote inteiro.
5. O sistema bloqueava cenários operacionais (conflito múltiplo) em vez de orientar o operador.

### Conclusão
O banco preservou a integridade. A falha foi **arquitetural na camada Service** — decisão operacional chegou à persistência sem plano validado.

---

## 2. Arquitetura Antes / Depois

### Antes
```
Lote → Análise (por lote) → Decisão UI → INSERT direto → IntegrityError → lote abortado
```

### Depois
```
Lote → Classificação por NF → Plano Operacional → Validação preventiva
     → Confirmação do operador (UI existente) → Execução → Auditoria → Relatório parcial
```

---

## 3. Arquivos Alterados

| Arquivo | Ação |
|---------|------|
| `carregamentos/services/nf_operacional_classifier.py` | **Criado** — máquina de classificação por NF |
| `carregamentos/services/validacao_item_carregamento.py` | **Criado** — validação preventiva alinhada à UNIQUE |
| `carregamentos/services/lote_processamento_service.py` | **Criado** — plano + execução resiliente |
| `carregamentos/services/decisao_operacional_auditoria_service.py` | **Criado** — rastreabilidade de decisões |
| `carregamentos/services/fechamento_service.py` | **Alterado** — delega complementação ao lote service |
| `carregamentos/services/analise_operacional_service.py` | **Alterado** — conflito/cancelada orientam, não bloqueiam automaticamente |
| `carregamentos/models/operacional.py` | **Alterado** — tipos do plano operacional |
| `carregamentos/models/fechamento.py` | **Alterado** — `relatorio_lote` no resultado |
| `carregamentos/integration.py` | **Alterado** — confirmação operacional para conflito |
| `infrastructure/models/constants.py` | **Alterado** — `AUDIT_EVENTO_DECISAO_OPERACIONAL` |
| `tests/test_reentrega_lote_resiliente.py` | **Criado** — cenários do incidente e auditoria |

**Não alterados:** layout UI, modelos ORM, migrations, UNIQUE constraints, repositories.

---

## 4. Regra de Negócio Implementada

### Classificação por NF (antes da persistência)
- `NOVA` → INSERT validado
- `REENTREGA` / `DUPLICIDADE` → reutilizar registros, sem INSERT
- `REIMPRESSAO` → PDF sem alterar itens
- `INVALIDA` → ocorrência registrada, lote continua

### Princípio operacional
- Sistema **diagnostica, orienta e recomenda**.
- Operador **decide** (confirmação via fluxos UI existentes).
- Bloqueio automático **somente** para falhas estruturais (`ErroEstruturalProcessamento`).

### Auditoria de decisão
Todo fechamento de complementação registra `DECISAO_OPERACIONAL` com:
usuário, data, hora, estação, motivo, decisão, situação anterior/posterior, NFs, impactos, riscos, recomendação.

---

## 5. Impactos

| Área | Impacto |
|------|---------|
| Complementação mista | NF com histórico não interrompe o lote |
| Reentrega / reimpressão | Sem regressão — fluxos preservados |
| Conflito múltiplo | Orientação + confirmação em vez de bloqueio automático |
| Performance | Validação adicional por NF (marginal) |
| UI | Sem alteração de layout — mensagens enriquecidas via `message`/`relatorio_lote` |

---

## 6. Riscos Residuais

| Risco | Mitigação |
|-------|-----------|
| Classificação incorreta por cache desatualizado | `invalidate_analise_operacional_cache()` após commit |
| Operador confirma sem ler diagnóstico | Auditoria completa + histórico operacional |
| Cenários edge não cobertos | Suite `test_reentrega_lote_resiliente.py` + 130 testes passando |

---

## 7. Plano de Rollback

1. Reverter commit desta correção (`git revert`).
2. Banco inalterado — sem migrations.
3. Comportamento retorna ao anterior (com risco de `IntegrityError` em lotes mistos).

---

## 8. Evidências de Testes

```
pytest tests/test_reentrega_lote_resiliente.py  → 5 passed
pytest (suite completa)                         → 130 passed, 3 skipped
```

### Cenários validados
- Lote misto com NF de reentrega (reprodução do incidente) — **0 IntegrityError**
- Complementação somente com reentregas — sem INSERT indevido
- Classificador marca duplicidade no mesmo carregamento
- Validação preventiva de duplicidade lógica
- Auditoria `DECISAO_OPERACIONAL` persistida
- Regressão: fechamento, análise operacional, fluxo histórico

---

## 9. Critérios de Aceite

| Critério | Status |
|----------|--------|
| Reentrega não interrompe lote | ✓ |
| Reimpressão sem duplicidade | ✓ |
| Complementação sem IntegrityError | ✓ |
| Decisão antes da persistência | ✓ |
| UNIQUE não usada como regra de negócio | ✓ |
| Operador com decisão final | ✓ |
| Falhas estruturais isoladas | ✓ |
| Funcionalidades homologadas compatíveis | ✓ |
| Testes automatizados aprovados | ✓ (130/130) |
