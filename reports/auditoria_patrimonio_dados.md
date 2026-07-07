# Homologacao do Banco de Producao

**Banco oficial:** `C:\MinutaData\minuta_dev.db`
**Localizado:** SIM
**Homologado:** SIM
**Apto migracao:** NAO
**Gerado em:** 2026-07-07T21:59:12.475137+00:00

## Resumo
- **banco_oficial:** C:\MinutaData\minuta_dev.db
- **banco_localizado:** True
- **banco_homologado:** True
- **apto_migracao:** False
- **tabelas:** 16
- **total_registros:** 5456
- **documento_xml:** 1046
- **documento_pdf:** 55
- **xml_fisicos:** 1046
- **pdf_fisicos:** 56
- **espaco_total_mb:** 17.5009
- **conclusao:** Banco oficial localizado e homologado com ressalvas. Revisar inconsistencias antes da migracao.
- **bloqueador:** Inconsistencias: pdf_ausente=5

## Inventario de Tabelas

| Tabela | Registros | Primeiro ID | Ultimo ID | MB est. |
|--------|-----------|-------------|-----------|---------|
| alembic_version | 1 | None | None | 0.0003 |
| carregamento | 30 | 1 | 30 | 0.0086 |
| configuracao | 1 | 1 | 1 | 0.0003 |
| destinatario | 0 | None | None | 0.0 |
| documento | 55 | 1 | 55 | 0.0157 |
| documento_xml | 1046 | 1 | 1046 | 0.2993 |
| evento_auditoria | 55 | 1 | 55 | 0.0157 |
| historico_operacional | 30 | 1 | 30 | 0.0086 |
| item_carregamento | 972 | 1 | 972 | 0.2781 |
| item_nota_fiscal | 2207 | 1 | 2207 | 0.6314 |
| motorista | 0 | None | None | 0.0 |
| nota_fiscal | 1055 | 1 | 1055 | 0.3018 |
| perfil | 2 | 1 | 2 | 0.0006 |
| rota | 0 | None | None | 0.0 |
| usuario | 2 | 1 | 2 | 0.0006 |
| veiculo | 0 | None | None | 0.0 |

## documento_xml x xml_storage

- Registros DB: 1046
- Arquivos FS: 1046
- OK: 1046
- Ausentes: 0
- FS sem registro: 0
- Hash divergente: 5

## documento x documentos

- Registros DB: 55
- Arquivos FS: 56
- OK: 50
- Ausentes: 5
- Hash divergente: 0

## Integridade

- fk_item_carregamento_orfa: 0
- fk_documento_carregamento_orfa: 0
- fk_item_nota_orfa: 0
- carregamentos_sem_itens: 0
- notas_sem_itens: 31
- historicos_orfaos: 0
- eventos_carregamento_orfaos: 0
- itens_sem_carregamento: 0

## Riscos

- **[BAIXO]** XML_BENCH_HASH: 5 XML(s) de benchmark com hash sintetico (bench-N).
- **[MEDIO]** NOTAS_SEM_ITENS: 31 nota(s)_fiscal sem item_nota_fiscal.
- **[ALTO]** PDF_AUSENTE: 5 PDF(s) no DB sem arquivo.
- **[ALTO]** HASH_DIVERGENTE: hash_xml=5, hash_pdf=0

## Conclusao

Banco oficial localizado e homologado com ressalvas. Revisar inconsistencias antes da migracao.