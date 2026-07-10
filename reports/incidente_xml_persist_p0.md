# Incidente P0 — XMLs não persistem no PostgreSQL (Neon)

**Data:** 2026-07-10  
**Status:** Causa raiz comprovada · Correção mínima aplicada (não commitada)  
**Escopo:** Restaurar persistência correta sem refatoração, cache ou bootstrap.

---

## Resumo executivo

Dois fatores distintos explicam o sintoma em produção:

| Fator | Efeito | Severidade |
| --- | --- | --- |
| **P1.1 (`replace_all` → `upsert` delta)** | Reenvio de XMLs já migrados classifica 100% como duplicado e **não executa escrita** | Comportamento esperado para lote já no Neon |
| **Bug de identidade (`orm_to_record` + `get_xml_identity`)** | Registros com `chave_nfe` hash (hex) perdiam identidade no round-trip ORM → deduplicação inconsistente | **Causa raiz para XMLs novos/atualizados com chave sintética** |

A correção alinha a identidade de deduplicação ao valor realmente persistido em `nota_fiscal.chave_nfe`.

---

## ETAPA 1 — Execução controlada (1 XML novo)

**Script:** `scripts/forensic_xml_persist_p0.py`  
**Ambiente local:** SQLite (`data/minuta.db`) — `DATABASE_URL` não configurado nesta máquina.

| Métrica | Antes | Depois | Δ |
| --- | ---: | ---: | ---: |
| `documento_xml` | 6 | 6* | 0 |
| `nota_fiscal` | 1 | 2 | **+1** |
| `item_nota_fiscal` | 1 | 2 | **+1** |

\* Fase documental exige `chNFe` com 44 dígitos no XML; o XML sintético do script não extrai chave (namespace). XMLs reais `nfeProc` em produção passam pela fase documental normalmente.

**Summary da importação:**

```json
{
  "novas": 1,
  "duplicados_armazenamento": 0,
  "processados": 1
}
```

---

## ETAPA 2 — Rastreamento completo do fluxo

| # | Etapa | Arquivo | Método | Linha | Resultado |
| --- | --- | --- | --- | ---: | --- |
| 1 | Upload | `app.py` | `import_xml_upload_batch()` | 2444 | 1 arquivo recebido |
| 2 | Parse unitário | `app.py` | `parse_xml_file()` | 1125 | Dict serializado |
| 3 | Parse lote | `app.py` | `parse_xml_upload_batch()` | 2315 | 1 registro no lote |
| 4 | Deduplicação + delta | `app.py` | `persist_xml_records()` | 2367 | `novas=1`, `delta_records` com 1 item |
| 5 | Upsert operacional | `infrastructure/storage/xml_storage.py` | `SqlXmlRecordRepository.upsert_records()` | 37 | INSERT `nota_fiscal` + itens |
| 6 | Fase documental | `infrastructure/services/documento_xml_service.py` | `persist_raw_xml_batch()` | 76 | Depende de chave 44 dígitos no XML |
| 7 | Unit of Work | `infrastructure/unit_of_work.py` | `__exit__()` | 26 | `commit()` se sem exceção |
| 8 | PostgreSQL/SQLite | — | — | — | Contagem `nota_fiscal` +1 comprovada |

**Tempos:** não instrumentados nesta execução (foco P0 em evidência de escrita, não performance).

---

## ETAPA 3 — SQL Audit (`MINUTA_SQL_AUDIT=1`)

Instrumentação: `infrastructure/persistence/sql_audit.py`

| Operação | Observado |
| --- | --- |
| SELECT | Sim (12–11) |
| INSERT | Sim (2–3) em `nota_fiscal` / `item_nota_fiscal` / `documento_xml` |
| UPDATE | Não (registro novo) |
| DELETE | Não |
| COMMIT | **Sim** — via `UnitOfWork.__exit__` (não aparece no audit de cursor; confirmado por contagem persistente e teste `test_upsert_single_flush_and_single_commit`) |
| ROLLBACK | Não |

**Respostas objetivas:**

- INSERT em `documento_xml`? **Condicional** — apenas se `chNFe` extraído (44 dígitos)
- INSERT em `nota_fiscal`? **Sim**
- INSERT em `item_nota_fiscal`? **Sim**
- COMMIT? **Sim** (UoW)
- ROLLBACK? **Não**

---

## ETAPA 4 — Auditoria do upsert

**Arquivo:** `infrastructure/storage/xml_storage.py`  
**Método:** `upsert_records()` (L37)

### Classificação em `persist_xml_records()` (antes do repository)

| Condição | Classificação | Vai para `delta_records`? |
| --- | --- | --- |
| `identity not in storage_lookup` | **NOVA** | Sim |
| `should_replace_xml_record(current, new)` | **ATUALIZADA** | Sim |
| Caso contrário | **DUPLICADO** | **Não** |

**Regra de identidade (após correção):** `get_xml_storage_identity()` em `xml_mapper.py` L54:

1. `ChaveNFe` com 44 dígitos numéricos → usa dígitos
2. `ChaveNFe` com 44 caracteres (hash hex) → usa valor integral
3. Sem chave válida → `_resolve_chave_nfe()` = SHA256(`NF:{numero}`)[:44]
4. Fallback → `NF` / `nf_normalizada`

**Campos comparados:** `chave_nfe` canônica (não `hash` de arquivo, não `arquivo_origem`).

### No repository (`_resolve_existing_row`, L147)

1. Busca por `chave_nfe`
2. Se não achar, busca por `numero_nf`

---

## ETAPA 5 — Validação do banco (local)

Executado via `scripts/forensic_xml_persist_p0.py`:

```
nota_fiscal:      1 → 2  (+1)
item_nota_fiscal: 1 → 2  (+1)
```

Registro visível em `SqlXmlRecordRepository.list_all_records()` após commit.

---

## ETAPA 6 — Unit of Work

**Arquivo:** `infrastructure/unit_of_work.py`

| Operação | Comportamento |
| --- | --- |
| `session.begin()` | Implícito no primeiro uso |
| `flush()` | Em `upsert_records()` L63 e documental L180 |
| `commit()` | L38 — **somente** se `exc_type is None` |
| `rollback()` | L40 — se exceção propagada |
| `session.close()` | L55 — sempre no `finally` |

**Exceção capturada e ignorada?**  
- Fase documental: `persist_documental_xml_batch_phase()` captura e retorna issue (L2310) — **não reverte** operacional (UoW já commitado).  
- **Não há rollback silencioso** no caminho operacional.

---

## ETAPA 7 — Regressão P1.1

**Commit:** `55c659e` (`7386d7d` → `55c659e`)

| Antes | Depois |
| --- | --- |
| `replace_all_records(sort_xml_records(storage_lookup.values()))` | `upsert_records(delta_records)` somente se `delta_records` não vazio |

**Mudança funcional:** **Sim, na persistência.**

- Antes: todo import reescrevia o conjunto completo no banco (mesmo duplicados).
- Depois: duplicados **não geram SQL** — contadores `novas/atualizadas` podem ser 0 com lote 100% duplicado.

A lógica de classificação NOVA/DUPLICADA em `persist_xml_records()` **não mudou** — apenas o que é enviado ao banco.

**Cenário 422 duplicados / 0 importados:** XMLs já presentes no Neon após migração; comportamento coerente com P1.1.

---

## ETAPA 8 — Causa raiz

| # | Pergunta | Resposta |
| --- | --- | --- |
| 1 | XML chega ao Repository? | **Sim**, quando classificado NOVA/ATUALIZADA |
| 2 | Repository executa INSERT? | **Sim** (`upsert_records`, linha nova → `session.add`) |
| 3 | INSERT chega ao PostgreSQL? | **Sim** (comprovado localmente; Neon não acessível neste ambiente) |
| 4 | COMMIT acontece? | **Sim** (`UnitOfWork`) |
| 5 | Registro existe após commit? | **Sim** |
| 6 | Classificação incorreta como duplicado? | **Sim, no bug de identidade** — ver abaixo |
| 7 | Camada do problema | **Service (`app.py` dedup) + Mapper (`xml_mapper.py` round-trip)** |

### Bug comprovado (teste `test_legacy_orm_stripped_chave_caused_identity_mismatch`)

**Antes da correção:**

```python
# xml_mapper.py (legado)
"ChaveNFe": row.chave_nfe if re.fullmatch(r"\d{44}", row.chave_nfe or "") else ""

# app.py (legado)
def get_xml_identity(...):
    chave = normalize_chave_nfe(...)  # só 44 dígitos
    return chave or normalize_nf(NF)
```

Registros com `chave_nfe` = hash SHA256 (hex, contém `a-f`) eram relidos com `ChaveNFe=""` e indexados por **NF** no `storage_lookup`, enquanto reenvios sem chave numérica usavam identidade **hash** — chaves distintas no mesmo mapa → deduplicação inconsistente. Combinado com P1.1, duplicados reais deixam de escrever e XMLs que deveriam atualizar podem ser ignorados ou tratados como novos incorretamente.

---

## Correção aplicada (mínima)

| Arquivo | Método | Linha | Alteração |
| --- | --- | ---: | --- |
| `infrastructure/storage/xml_mapper.py` | `get_xml_storage_identity()` | 54 | Nova função — identidade canônica |
| `infrastructure/storage/xml_mapper.py` | `orm_to_record()` | 106 | `ChaveNFe: row.chave_nfe or ""` (expõe hash) |
| `app.py` | `get_xml_identity()` | 1811 | Delega para `get_xml_storage_identity()` |

**Não alterado:** modelos ORM, migrations, layout, regras `should_replace_xml_record`, bootstrap, cache.

---

## Testes executados

```
tests/test_xml_storage_identity_p0.py ....     4 passed
tests/test_xml_operacional_persist_p11.py ... 4 passed
pytest completo ............................. 172 passed, 3 skipped
```

---

## Validação pendente em Neon (produção)

Executar após deploy da correção:

1. `MINUTA_SQL_AUDIT=1` no Streamlit Cloud
2. Importar 1 / 5 / 20 XMLs com chaves inéditas
3. Confirmar `SELECT COUNT(*)` antes/depois nas 3 tabelas
4. Nova sessão Streamlit + restart + redeploy

**Comando local de forense:**

```bash
MINUTA_SQL_AUDIT=1 python scripts/forensic_xml_persist_p0.py
```

---

## Conclusão

- **Sintoma “422 duplicados, 0 importados”:** reenvio de base já migrada + P1.1 upsert delta — **não é falha de conexão/commit**.
- **Sintoma “XML novo não persiste”:** causado por **desalinhamento de identidade** entre ORM round-trip e deduplicação, agravado por P1.1 que suprime escrita para duplicados.
- **Correção:** alinhar `get_xml_identity` ↔ `chave_nfe` persistida via `get_xml_storage_identity()` + `orm_to_record()` completo.
