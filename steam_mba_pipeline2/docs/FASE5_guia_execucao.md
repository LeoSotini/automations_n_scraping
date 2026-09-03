# FASE 5 — Guia de Execução Passo a Passo

Guia operacional para você executar a coleta completa. Escrito para ser seguido
sem consultar mais nada.

**Tudo é interrompível e retomável.** Você nunca perde trabalho ao fechar o
terminal, reiniciar a máquina ou apertar `Ctrl+C`.

---

## 0. Antes de começar

### 0.1. Abra um terminal na pasta do projeto

```powershell
cd "c:\Users\z004kree\OneDrive - Siemens Energy\Área de Trabalho\MBA DSA\steam-req-pipeline"
```

### 0.2. Confirme que o ambiente está íntegro

```powershell
python -m pytest -q
```

Esperado: `160 passed`. Se algum teste falhar, **não prossiga** — algo mudou no
ambiente ou na estrutura da Steam.

### 0.3. Confirme o estado atual

```powershell
python main.py status
```

Você verá o que já foi coletado no piloto (~310 jogos). Esses registros **serão
reaproveitados**: a coleta completa os pula sem gastar requisição.

### 0.4. Revise o ritmo configurado

Abra `config\settings.yaml` e confira:

```yaml
scrape:
  sleep_between_apps_s: 2.0     # decidido no piloto (D12): 700 reqs, 0 429
```

Aumentar esse número deixa a coleta mais lenta e mais conservadora. Alterar não
exige mexer em código nem invalida o que já foi coletado.

---

## 1. Discovery — enumerar a população

**Duração estimada: ~5,5 h** (2.540 páginas do discovery + 2.380 do ledger de
Indie, a 4,0 s por requisição).

```powershell
python main.py discover
```

### O que você verá

```
[TAGS]      446 tags oficiais | 42/42 tags do filtro resolvidas
[DISCOVERY] 24 tags positivas a consultar | untags=Indie(492) | ritmo 4.0s/req
[DISCOVERY] --- tag 1/24: 3D ---
[DISCOVERY] tag 3D (4191) untags=Indie(492) -> total_count=26,755
[DISCOVERY] 3D             | pag 0005/0536 | 250 app_ids | 0.9% | decorrido 00:00:20
...
[DISCOVERY] 3D             | cobertura 99.2% (26,542 distintos de 26,755 declarados)
[DISCOVERY] 3D             | +26,542 novos | uniao=26,542
```

### O que observar

- **`cobertura`** ao fim de cada tag. Acima de 95% está bom. Se aparecer
  `WARNING` de cobertura baixa, rode `python main.py discover` de novo depois —
  a união é cumulativa e deduplicada, então a segunda passada recupera o que
  faltou.
- **`total_count`** por tag. Se `3D` vier muito abaixo de ~26.000, algo mudou no
  endpoint e vale me avisar antes de seguir.
- **`Rail Shooter`** vai reportar `total_count=0` com um WARNING. **Isso é
  esperado** — a tag existe na taxonomia mas não tem jogos.

### Interromper e retomar

`Ctrl+C` a qualquer momento. Para continuar:

```powershell
python main.py discover          # retoma automaticamente
```

O estado fica em `data\checkpoints\discovery_state.json`, por tag e por página.

### Saídas

| Arquivo | Conteúdo |
|---|---|
| `data\processed\candidates.json` | a população a raspar |
| `data\processed\ledger_excluded_indie.json` | os app_ids excluídos por Indie (item 2.17) |

### Se quiser pular o ledger de Indie

Ele consome ~2,6 h e serve para manter a decisão D2 reversível. Para pular:

```powershell
python main.py discover --skip-indie-ledger
```

Recomendo **não pular**: sem ele, revisar o critério de Indie no futuro exigiria
repetir o discovery inteiro.

---

## 2. Scrape — coletar requisitos e tags

**Duração estimada a 2 s/jogo: 44 h a 89 h**, conforme a sobreposição real entre
tags. Você descobrirá o número exato ao fim do discovery
(`uniao de 24 tags positivas -> N app_ids unicos`).

### 2.1. Rode primeiro um teste curto

Antes de largar por dias, confirme que tudo está saudável:

```powershell
python main.py scrape --limit 50
```

Cheque no output: `0 falhas` e `0 rate-limited`. Se aparecerem 429, aumente
`sleep_between_apps_s` para 4.0 e repita.

### 2.2. Rode a coleta completa

```powershell
python main.py scrape
```

### O que você verá

```
[SCRAPE]    63,510 app_ids na fila | tags do HTML=True | ritmo appdetails=0.6s | pausa entre jogos=2.0s
[SCRAPE]    estimativa de duracao: 63.5 h (2.6 dias) no ritmo configurado
[SCRAPE 00431/63510 | 0.7%] app_id=1091500 | Cyberpunk 2077
   [OK] appdetails 34.3 KB | type=game | release=Dec 9, 2020 | reviews=886,479
   [OK] 20 tags via STORE_HTML_AFTER_AGE_GATE
   [OK] requisitos RECOMENDADOS capturados (formato A)
[CHECKPOINT] 431 completos | 12 parciais | 5 falhas | 3 pulados | 63,059 restantes
             estado global: COMPLETE=431 PARTIAL=12 FAILED=5 PENDING=63,062
             decorrido 00:23:12 | ETA 56:44:10 (~3.22s/app) | HTTP: 886 reqs, 0 retries, 0 rate-limited
```

Um `[CHECKPOINT]` a cada 25 jogos, com ETA recalculado.

### O que é normal e não requer ação

| Mensagem | Significado |
|---|---|
| `[INFO] sem requisitos recomendados na Steam` | o jogo realmente não os publica (~5%) |
| `[WARN] AGE_GATE_LOGIN_REQUIRED` | exige login; **não contornamos** por decisão. Os requisitos são coletados de todo modo; só as tags faltam |
| `[FAIL] appdetails success=false` | app removido ou indisponível na região |
| `[RATE] HTTP_429 -> backoff 4.0s` | a Steam pediu calma; o backoff já cuidou |
| `parciais` | jogo com requisitos coletados mas sem tags |

### O que merece atenção

| Mensagem | O que fazer |
|---|---|
| `[WARN] InitAppTagModal ausente e SEM age gate` | possível mudança na estrutura da página da Steam. Se aparecer **muitas vezes**, me avise |
| muitos `rate-limited` no `[CHECKPOINT]` | aumente `sleep_between_apps_s` e retome |
| `falhas` crescendo rápido | interrompa e verifique sua conexão |

### Interromper e retomar

```
Ctrl+C
```

O pipeline grava o checkpoint e mostra o resumo antes de sair. Para continuar:

```powershell
python main.py resume
```

Ele **pula tudo que já está pronto sem gastar requisição**, e a pausa de 2 s não
é aplicada aos pulados — retomar é rápido.

### Rodar em blocos ao longo de vários dias

Perfeitamente suportado. Rode algumas horas, `Ctrl+C`, e no dia seguinte
`python main.py resume`. Repita até `PENDING=0`.

Para limitar cada sessão:

```powershell
python main.py scrape --limit 5000     # processa 5.000 e para
```

---

## 3. Reprocessar as falhas

Depois que o scrape terminar:

```powershell
python main.py status          # veja quantos FAILED existem
python main.py retry-failed    # reprocessa apenas eles
```

Falhas por timeout ou rede tendem a passar na segunda tentativa. As que
persistirem com `APPDETAILS_SUCCESS_FALSE` são apps removidos da Steam — dado
ausente de verdade, não erro.

---

## 4. Filtro, validação e export

Estas três etapas são **offline**: leem `data\raw\` e **não fazem nenhuma
requisição**. Rodam em segundos e podem ser repetidas à vontade.

```powershell
python main.py filter
python main.py validate
python main.py export --flat
```

### 4.1. `filter` — o funil da metodologia

```
[FILTER]    entrada 63,510 registros (a partir de data/raw)
[FILTER]      NOT_GAME               -> -1,204  (62,306 restantes)
[FILTER]      BEFORE_2005            -> -2,118  (60,188 restantes)
[FILTER]      TWO_DIMENSIONAL        -> -4,003  (56,185 restantes)
[FILTER]    conflitos 3D+2D preservados: 1,915 (needs_manual_review=true)
[FILTER]    saida 56,185 registros incluidos
```

**Guarde este output** — é o funil que vai na metodologia do TCC.

### 4.2. `validate` — precisa terminar com 0 erros

```
[VALIDATE]  63,510 registros verificados | 0 erros | 214 avisos
[VALIDATE]  OK: nenhum erro bloqueante
```

Avisos são aceitáveis. `UNPARSED_LABELS` são rótulos exóticos preservados;
`INDIE_PASSED_SOURCE_FILTER` só deve aparecer para os jogos do piloto que foram
coletados por `--app-ids`. **Se aparecerem erros, me avise** — não use o dataset
sem investigar.

### 4.3. `export` — o dataset final

- `--flat` gera as colunas `minimum_cpu`, `recommended_cpu`, etc., prontas para
  Pandas/Polars. **Use `--flat` para a análise.**
- sem `--flat`, mantém `pc_requirements` aninhado.

| Arquivo | Conteúdo |
|---|---|
| `data\processed\dataset.json` | **o dataset** |
| `data\processed\dataset_metadata.json` | colunas, versões, funil, limitações |
| `data\processed\ledger_exclusions.json` | todo descarte com motivo |
| `data\processed\validation_report.json` | relatório de validação |

Carregar no Pandas:

```python
import pandas as pd
df = pd.read_json(r"data\processed\dataset.json")
print(df.shape)
print(df["recommended_ram"].notna().mean())
```

---

## 5. Threshold de reviews (item 2.14)

Ainda **não** aplicado, por decisão metodológica. Para produzir a evidência que
o item 2.14 exige antes de fixar o corte:

```powershell
python main.py analyze-thresholds
```

Saída: distribuição de `review_count` e quantos sobrevivem a 0 / 100 / 500 /
1.000 / 5.000. Com isso em mãos, escolha o valor, coloque em
`config\filters.yaml` (`reviews.min_threshold`), **incremente
`filter_version`** e rode `filter` + `export` de novo — sem re-raspar nada.

---

## 6. Auditoria do viés de sobrevivência (opcional, decisão D1)

Quantifica quantos jogos do catálogo não aparecem no discovery — a limitação
mais séria a declarar na metodologia.

1. Gere uma chave gratuita em https://steamcommunity.com/dev/apikey
2. No terminal:

```powershell
$env:STEAM_API_KEY = "sua_chave_aqui"
python main.py audit-bias
```

Sem a chave, o comando apenas avisa e sai — nada quebra. Resultado em
`data\processed\bias_audit.json`, estratificado por ano de lançamento.

---

## 7. Resumo dos comandos

| Comando | Duração | Rede? |
|---|---|---|
| `python main.py tags` | segundos | sim |
| `python main.py discover` | ~5,5 h | sim |
| `python main.py scrape` | 44–89 h | sim |
| `python main.py resume` | o que restar | sim |
| `python main.py retry-failed` | minutos | sim |
| `python main.py filter` | segundos | **não** |
| `python main.py validate` | segundos | **não** |
| `python main.py export --flat` | segundos | **não** |
| `python main.py analyze-thresholds` | segundos | **não** |
| `python main.py audit-bias` | ~10 min | sim |
| `python main.py status` | instantâneo | **não** |

---

## 8. Solução de problemas

**"candidates.json vazio ou ausente"** — rode `discover` antes do `scrape`.

**Muitos HTTP 429** — aumente `sleep_between_apps_s` em `settings.yaml` (tente
4.0 ou 6.0) e rode `resume`. Nada do que já foi coletado se perde.

**O processo morreu sem aviso (queda de energia, reinício)** — rode
`python main.py status` e depois `resume`. A escrita atômica do checkpoint
garante que o arquivo nunca fica corrompido.

**Erro `ERRO DE CONFIGURACAO: tags inexistentes na taxonomia`** — a Steam
renomeou ou removeu uma tag. Corrija o nome em `config\filters.yaml`. Esse abort
é intencional: prosseguir produziria um filtro silenciosamente vazio.

**Quero recomeçar o discovery do zero** — `python main.py discover --no-resume`.
Não apague `data\raw\`: são os payloads brutos, e perdê-los significa re-raspar
tudo.

**Quero refazer o filtro com outro critério** — edite `config\filters.yaml`,
**incremente `filter_version`**, e rode `filter` + `validate` + `export`. Zero
requisições à Steam. Esta é a propriedade central do desenho: os critérios
metodológicos são revisáveis sem recoletar.

**A pasta `data\raw\` está grande** — é esperado (~2 payloads por jogo). É a
fonte de verdade do projeto; mantenha, e de preferência faça backup dela junto
com `config\`.

---

## 9. O que me reportar ao terminar

Para que eu possa validar a coleta e ajudar na próxima etapa:

1. o output completo do `filter` (o funil);
2. as contagens do `validate` (erros e avisos);
3. `data\processed\dataset_metadata.json`;
4. quantos `FAILED` restaram após o `retry-failed`;
5. qualquer WARNING que tenha aparecido com frequência.
