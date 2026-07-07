# Encerramento FASE N4 — Migração SQLite → PostgreSQL Neon

**Status:** CONCLUÍDO  
**Tag oficial:** `v2.0.0-postgresql`  
**Data:** 2026-07-07

---

## Declaração formal

As fases **P0, N1, N2, N3 e N4** estão encerradas.

O **PostgreSQL Neon** passa a ser o banco de dados principal e oficial do sistema.

O **SQLite** (`C:\MinutaData\minuta_dev.db`) permanece apenas como contingência e backup.

O projeto está **apto para iniciar a Fase P1** — Otimização de Performance e Estabilização Pós-Migração.

---

## Validações executadas

| Item | Resultado |
|------|-----------|
| PostgreSQL conectado | OK |
| Engine SQLAlchemy PostgreSQL | OK |
| MINUTA_DATABASE_URL → Neon | OK (.env local, não versionado) |
| Alembic HEAD | `m5_0005_operational_tables` |
| DatabaseUsageService | OK |
| Pytest completo | **125 passed**, 3 skipped |
| Homologação N4 (audit) | APROVADA |
| Environment diagnose | AMBIENTE PRONTO PARA MIGRACAO |
| Integridade pós-migração | equivalente = true |
| XML filesystem | 1.046/1.046 |
| Credenciais no Git | **Nenhuma** |

---

## Banco oficial

| Métrica | Valor |
|---------|------:|
| PostgreSQL | 18.4 (Neon) |
| Tabelas domínio | 16 |
| Registros migrados | 5.455 |
| Foreign keys | 20 |
| Índices | 100 |
| documento_xml | 1.046 |
| carregamento | 30 |
| nota_fiscal | 1.055 |
| Espaço Neon | ~12,2 MB (2,4%) |

---

## Segurança

- `.env` está em `.gitignore` — **não será commitado**
- Nenhuma senha (`npg_*`) em arquivos versionados
- Reports contêm apenas hosts mascarados (`***` na URL)

---

## Rollback

```
Backup: C:\MinutaData\backups\backup-20260707-181552
SQLite: C:\MinutaData\minuta_dev.db

Procedimento: MINUTA_DATABASE_URL=sqlite:///C:/MinutaData/minuta_dev.db + reiniciar app
```

---

## Riscos remanescentes

1. 5 PDFs ROMANEIO ausentes no disco (pré-existente)
2. Senha Neon exposta em sessão de deploy — **rotacionar no console Neon**
3. Tabela `minuta_neon_probe` (homologação) com 1 registro residual

---

## Recomendações futuras (Fase P1)

- Otimizar `list_all()` em validação NF
- Índices parciais `documento_xml WHERE ativo`
- Monitoramento de queries e bloqueio emergência 95%+
- Rotacionar credenciais Neon
