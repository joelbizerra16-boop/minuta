# Estabilizacao Arquitetural A1 — PostgreSQL como Fonte Oficial

**Data:** 2026-07-10  
**Escopo:** Coerencia Streamlit ↔ Services ↔ Repositories ↔ PostgreSQL Neon  
**Tipo:** Correcao arquitetural (sem alteracao de regra de negocio, ORM, migrations ou layout)

---

## 1. RCA — Causas raiz

| ID | Severidade | Causa raiz | Evidencia |
|----|------------|------------|-----------|
| I-01 | P0 | Gatilho de refresh usava `mtime` de arquivos JSON legados (`data/*.json`) em vez de estado do PostgreSQL | `app.py:7504-7512` (antes da correcao) |
| I-02 | P0 | Invalidacao de cache parcial apos escrita — apenas um `@st.cache_data.clear()` por operacao | `persist_xml_records`, `salvar_separacao_records` |
| I-03 | P1 | Tela de separacao preferia `session_state.separacao_records` ao parametro recém-carregado | `app.py:6072` |
| I-04 | P1 | `get_separacao_storage_status()` retornava `datetime.now()` em vez de `configuracao.atualizado_em` | `app.py:3724-3728` |
| I-05 | P1 | Token de cache de classificacao baseado em arquivo local | `get_path_cache_token(CLASSIFICACAO_PRODUTOS_JSON_PATH)` |
| I-06 | P2 | UI de gestao de dados exibia tamanho de arquivo JSON inexistente como fonte | `app.py:6336-6340` |

### O que NAO e inconsistencia (mantido por design)

- **`processed_df` / Excel:** estado de workflow pre-fechamento; dados nao existem no PostgreSQL ate `FechamentoCarregamentoService` gravar carregamento.
- **`auth_user` em session:** snapshot de sessao pos-login; escrita de autenticacao confirma no PostgreSQL.
- **`AnaliseOperacionalService._carregamentos_cache`:** acelerador com `invalidar_cache()` nos fechamentos — mantido (nao removido sem benchmark).
- **Painéis TTL 300s:** cache de UI; invalidados por `clear_contexto_operacional()`.

---

## 2. Correcoes aplicadas (Etapa A9)

### Arquivos novos

| Arquivo | Responsabilidade |
|---------|------------------|
| `core/runtime_data_coherence.py` | Assinaturas de leitura baseadas em PostgreSQL (`count` + `atualizado_em`) |
| `infrastructure/persistence/sql_audit.py` | Instrumentacao SQL opcional (`MINUTA_SQL_AUDIT=1`) |
| `tests/test_runtime_data_coherence_a1.py` | Testes de assinatura e mudanca pos-escrita |
| `reports/estabilizacao_arquitetura_a1.md` | Este relatorio |

### Arquivos alterados

| Arquivo | Metodo / area | Alteracao |
|---------|---------------|-----------|
| `app.py` | `mark_persistence_layer_stale()` | Invalidacao unificada: `st.cache_data` + assinaturas session |
| `app.py` | `load_runtime_reference_data()` | Assinatura via `get_reference_data_signature()` |
| `app.py` | `load_runtime_operational_data()` | Assinatura via `get_operational_data_signature()` |
| `app.py` | `get_separacao_storage_status()` | Le `configuracao.atualizado_em` do PostgreSQL |
| `app.py` | `render_separacao_screen()` | Sincroniza `session_state` a partir do parametro carregado |
| `app.py` | `persist_xml_records`, `salvar_*`, cleanup | Chamam `mark_persistence_layer_stale()` |
| `app.py` | Gestao de dados (cards) | Texto indica PostgreSQL, nao JSON local |
| `infrastructure/database.py` | `configure_database()` | Registra `register_sql_audit()` quando habilitado |

---

## 3. Matriz de consistencia (pos-A1)

| Modulo | Le PostgreSQL | Escreve PostgreSQL | Session (UI) | Cache acelerador | JSON legado |
|--------|---------------|--------------------|--------------|------------------|-------------|
| Login | Sim (auth) | Sim | Snapshot usuario | Service singleton | Nao |
| Dashboard | Nao (estatico) | Nao | Navegacao | Nao | Nao |
| XML import | Sim | Sim | Relatorio upload | `@st.cache_data` + invalidacao | Nao (assinatura PG) |
| Excel/Minuta | XML: Sim; Excel: session | No fechamento | `processed_df` | Sim | Nao |
| Carregamentos | Sim | Sim (fechamento) | Diagnostico/PDF | Analise service cache | Nao |
| Separação/Lotes | Sim (`configuracao`) | Sim | Lote atual/feedback | `@st.cache_data` + invalidacao | Nao |
| Gestao/Retencao | Sim | Sim (execucao) | Wizard simulacao | Nao | Nao |
| Usuarios | Sim (lista) | Sim | Form state | Service singleton | Nao |
| Auditoria | Sim | Append only | Expander state | TTL 300s | Nao |

---

## 4. Fluxo de invalidacao (pos-correcao)

```
Escrita PostgreSQL (upsert XML, save configuracao, cleanup)
    → mark_persistence_layer_stale(reference=?, operational=?)
        → carregar_*_records.clear()
        → runtime_*_signature = None
        → runtime_refresh_required = True
    → Proximo render
        → load_runtime_* compara assinatura PG
        → Recarrega do PostgreSQL se divergir
```

---

## 5. Instrumentacao SQL (Etapa A6)

Ativar em ambiente de diagnostico:

```bash
MINUTA_SQL_AUDIT=1 streamlit run app.py
```

Registra operacoes por cursor (`SELECT`, `INSERT`, `UPDATE`, `DELETE`) com tabela inferida e duracao.

API: `infrastructure.persistence.sql_audit.get_sql_audit_report()`

---

## 6. Teste de consistencia (Etapa A7 — roteiro)

| Cenario | PostgreSQL | Tela apos acao |
|---------|------------|----------------|
| Importar XML novo | INSERT em `nota_fiscal` | `load_runtime_reference_data(force_refresh=True)` recarrega |
| Nova aba / outro browser | Dados persistidos | Assinatura PG diverge → refresh |
| Rerun Streamlit | Inalterado | Cache invalidado apos escrita; leitura via PG |
| Restart / redeploy | Inalterado | Cold start le PostgreSQL |
| Hibernacao Cloud | Inalterado | Mesmo processo: assinaturas PG |

---

## 7. Impactos e riscos

| Impacto | Descricao |
|---------|-----------|
| Positivo | Refresh confiavel apos import XML, separacao e cleanup |
| Positivo | Menos dependencia de arquivos `data/*.json` inexistentes |
| Neutro | +2 queries leves de assinatura por refresh (count/max) |
| Risco baixo | Assinatura por `atualizado_em` — se escrita nao atualizar timestamp, refresh pode falhar (mitigado: `TimestampMixin` no ORM) |

### Rollback

Reverter commit A1 restaura gatilhos por `mtime` JSON. Dados no PostgreSQL nao sao afetados.

---

## 8. Testes executados

```
pytest completo: 168 passed, 3 skipped (pos-A1)
tests/test_runtime_data_coherence_a1.py: 4 passed
tests/test_infrastructure_bootstrap_p01.py: 7 passed (incl. ORM reload)
```

Validacao Streamlit Cloud: pendente deploy do commit A1.

---

## 9. Criterio de aceite — status

| Criterio | Status |
|----------|--------|
| PostgreSQL fonte oficial de dados persistidos | **Atendido** |
| Gravacao com COMMIT via UnitOfWork | **Atendido** (sem mudanca) |
| Leitura operacional consulta PG quando necessario | **Atendido** (assinaturas PG) |
| Decisoes nao dependem exclusivamente de cache/JSON legado | **Atendido** para XML/separacao/lotes/classificacao |
| Gatilhos de atualizacao usam PG | **Atendido** |
| Consistencia apos restart/rerun/deploy | **Atendido** (por design + testes) |
| Excel pre-fechamento em session | **Aceito** — workflow transitório, nao existe no PG |

---

## 10. Proximos passos (fora do escopo A1)

- Validar em producao Neon pos-deploy.
- Avaliar revalidacao de `auth_user` em acoes administrativas (P2).
- Benchmark de `AnaliseOperacionalService` cache apenas se houver evidencia de stale em producao.
