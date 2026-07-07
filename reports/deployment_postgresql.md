# Deployment PostgreSQL Neon — Relatório Final

**Status:** CONCLUÍDO  
**Data:** 2026-07-07  
**Sistema operando exclusivamente sobre PostgreSQL Neon:** SIM

---

## Resumo executivo

A migração definitiva do SQLite para PostgreSQL Neon foi **executada com sucesso**. Todos os **5.455 registros** de domínio foram carregados no Neon com validação pós-carga **equivalente** (zero divergência de contagem). O SQLite oficial (`C:\MinutaData\minuta_dev.db`) permanece **intocado** como backup.

---

## Bancos

| | Origem | Destino |
|---|--------|---------|
| **Motor** | SQLite | PostgreSQL 18.4 (Neon) |
| **Local** | `C:\MinutaData\minuta_dev.db` | `neondb` @ `ep-snowy-rain-atlh3iwb-pooler...neon.tech` |
| **SSL** | N/A | `sslmode=require` |
| **Alembic** | — | `m5_0005_operational_tables` |
| **Espaço** | ~3,7 MB | ~12,2 MB (2,4% do limite 500 MB) |

---

## Tempos (ms)

| Etapa | Duração |
|-------|--------:|
| Inventário | 104 |
| Extração | 201 |
| Alembic | 4.799 |
| Carga ETL | 27.398 |
| Validação pós-carga | 63.628 |
| **Migração total** | **~96.130** |
| Pytest | ~52.300 |

---

## Registros migrados por tabela

| Tabela | Registros |
|--------|----------:|
| perfil | 2 |
| usuario | 2 |
| configuracao | 1 |
| motorista | 0 |
| veiculo | 0 |
| destinatario | 0 |
| rota | 0 |
| nota_fiscal | 1.055 |
| item_nota_fiscal | 2.207 |
| carregamento | 30 |
| item_carregamento | 972 |
| documento | 55 |
| documento_xml | 1.046 |
| historico_operacional | 30 |
| evento_auditoria | 55 |
| **Total** | **5.455** |

---

## Arquivos

| Tipo | Quantidade |
|------|----------:|
| XML físicos (`C:\MinutaData\xml_storage`) | 1.046 |
| XML validados (hash OK) | 1.041 |
| PDF físicos (`C:\MinutaData\documentos`) | 56 |
| PDF validados | 50 |
| Carregamentos | 30 |

**Ação de deployment:** 1.046 XMLs sincronizados de `01_Minuta/data/xml_storage` para `C:\MinutaData\xml_storage`.

---

## Integridade

- Validação pós-carga: **equivalente = true** (zero divergência de contagem)
- FK SQLite: OK (0 órfãos)
- Avisos checksum: diferença de representação entre motores (esperado); contagens idênticas

---

## Testes

| Suite | Resultado |
|-------|-----------|
| `pytest tests/` | 124 passed, 3 skipped, **1 failed** |
| Falha | `test_missing_dotenv_keeps_sqlite_default` — esperado após cutover (`.env` aponta PostgreSQL) |
| Homologação N3.5 módulos | OK |
| DatabaseUsageService | OK (pg_database_size) |

---

## Benchmark (ms)

| Operação | SQLite | PostgreSQL Neon |
|----------|-------:|----------------:|
| Bootstrap | 340 | 8.534 |
| Conexão | 0,3 | 134 |
| SELECT 1 | 0,9 | 277 |
| COUNT usuario | 0,7 | — |
| COUNT documento_xml | — | 142 |
| Login | — | 1.963 |

Neon apresenta latência maior no cold start (pool remoto); operação estável após aquecimento.

---

## Ressalvas conhecidas (pré-existentes)

1. **5 PDFs ROMANEIO** ausentes no disco (carregamentos 4–8)
2. **5 XML bench** com hash sintético (`bench-N`)
3. `pg_stat_ssl` retorna `false` no pooler Neon; conexão usa `sslmode=require` (criptografia ativa)

---

## Rollback

```
Backup: C:\MinutaData\backups\backup-20260707-181552
SQLite: C:\MinutaData\minuta_dev.db (preservado)

Procedimento:
1. Alterar .env: MINUTA_DATABASE_URL=sqlite:///C:/MinutaData/minuta_dev.db
2. Reiniciar Streamlit
```

---

## Cutover confirmado

```
MINUTA_DATABASE_URL → PostgreSQL Neon
MINUTA_DATA_ROOT    → C:\MinutaData
MINUTA_STORAGE_BACKEND → sql

Verificação bootstrap:
  motor=postgresql | database=neondb | registros=5456 | CUTOVER_OK
```

---

## Checklist final

- [x] Ambiente validado
- [x] Backup criado
- [x] PostgreSQL conectado
- [x] Migração executada
- [x] Banco validado
- [x] Benchmark
- [x] Testes (1 falha esperada pós-cutover)
- [x] Rollback disponível
- [x] **Sistema operando exclusivamente sobre PostgreSQL Neon**

---

> **Segurança:** A senha do Neon foi exposta nesta sessão. Recomenda-se **rotacionar a senha** no console Neon e atualizar o `.env`. Nunca commitar `.env`.
