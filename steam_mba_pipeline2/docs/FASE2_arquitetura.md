# FASE 2 — Especificação de Arquitetura

**Versão da especificação:** 1.0
**Baseada em:** `docs/FASE1_investigacao.md` (sondagem empírica de 2026-08-30)
**Decisões metodológicas:** registradas em `docs/DECISOES.md`

---

## 1. Decisões que governam esta arquitetura

| # | Decisão | Implicação arquitetural |
|---|---|---|
| **Q1** | Discovery **híbrido**: search como caminho principal + Steam Web API key para **auditar** o viés de sobrevivência | Dois módulos de discovery independentes; o de auditoria é opcional e não bloqueia o pipeline |
| **Q2** | `Indie` **excluído na fonte** (`untags=492`) | Discovery mais barato; **exige ledger dos excluídos** para reversibilidade (§4.3) |
| **Q3** | Tags a partir do **HTML de cada jogo** (20 tags via `InitAppTagModal`) | 2 requisições por candidato; estágio de enriquecimento separado e retomável |
| **Q4** | `appdetails.recommendations.total` é a métrica **oficial** de reviews | Zero requisições extra; `appreviews` **não** será chamado |

### 1.1. Tensão declarada entre Q2 e o §2.16

O §2.16 do projeto prioriza recall; a decisão Q2 exclui 48,4% da população 3D na
fonte. A tensão é real e será **mitigada, não escondida**:

- os app_ids excluídos por `Indie` serão gravados em
  `data/processed/ledger_excluded_indie.json` com `exclusion_reason=INDIE` e
  `filter_version`;
- o ledger custa apenas requisições de *discovery* (sem `appdetails`), pois só
  precisa de app_id + nome;
- revisar o §2.10 no futuro passa a exigir apenas o estágio de scraping sobre o
  ledger — **sem repetir o discovery**;
- a limitação será declarada explicitamente na metodologia do TCC.

---

## 2. Arquitetura

### 2.1. Estrutura de diretórios

```
steam-req-pipeline/
├── config/
│   ├── settings.yaml            # rate limits, timeouts, retries, caminhos
│   └── filters.yaml             # metodologia §2 versionada (filter_version)
├── data/
│   ├── raw/                     # payloads brutos, imutáveis
│   │   ├── appdetails/          # <app_id>.json (resposta íntegra)
│   │   └── store_html/          # <app_id>.tags.json (só as tags extraídas)
│   ├── processed/
│   │   ├── candidates.json      # saída do discovery+filtro
│   │   ├── ledger_excluded_indie.json
│   │   ├── ledger_exclusions.json   # §2.17: todo descarte com motivo
│   │   └── dataset.json         # DATASET FINAL (JSON tabular)
│   ├── checkpoints/
│   │   ├── discovery_state.json
│   │   └── scrape_state.json    # estado por app_id
│   └── logs/
│       └── run_YYYYMMDD_HHMMSS.log
├── docs/
│   ├── FASE1_investigacao.md
│   ├── FASE2_arquitetura.md
│   └── DECISOES.md
├── investigation/               # scripts da FASE 1 (reprodutibilidade)
├── src/steamreq/
│   ├── __init__.py
│   ├── config.py                # carga e validação de YAML
│   ├── logging_setup.py         # console verbose + arquivo
│   ├── steam_client.py          # HTTP: rate limit, retry, backoff, 429/5xx
│   ├── tags.py                  # taxonomia oficial; resolve nome -> tagid
│   ├── discovery.py             # search paginado -> candidatos
│   ├── discovery_audit.py       # opcional: API key, mede viés (Q1)
│   ├── filters.py               # metodologia §2, com exclusion_reason
│   ├── requirements_parser.py   # pc_requirements -> campos estruturados
│   ├── scraper.py               # orquestra appdetails + tags por app_id
│   ├── storage.py               # escrita atômica, checkpoints, ledgers
│   ├── validators.py            # §8 do projeto
│   └── export.py                # dataset.json tabular + metadados
├── tests/
├── main.py                      # CLI
├── pyproject.toml
└── README.md
```

### 2.2. Fluxo de dados

```
                    ┌──────────────────────────────────┐
                    │ IStoreService/GetTagList/v1      │
                    │ 446 tags oficiais (sem auth)     │
                    └────────────────┬─────────────────┘
                                     │ resolve nome -> tagid
                                     │ FALHA RUIDOSA se tag inexistir
                                     ▼
  filters.yaml ──────────►  ┌──────────────────────────┐
  (§2.5-2.13)               │  ESTÁGIO 1: DISCOVERY    │
                            │  /search/results/        │
                            │  category1=998 (jogos)   │
                            │  os=win                  │
                            │  tags=<T> untags=492     │
                            │  1 query POR tag positiva│  <- search é AND!
                            │  união no cliente        │
                            └────────────┬─────────────┘
                                         │ app_id, nome, tagids parciais
                                         ▼
                            ┌──────────────────────────┐
                            │  candidates.json         │
                            │  + ledger_excluded_indie │
                            └────────────┬─────────────┘
                                         ▼
                            ┌──────────────────────────┐
                            │  ESTÁGIO 2: SCRAPE       │
                            │  por app_id:             │
                            │   a) appdetails (JSON)   │
                            │   b) HTML -> 20 tags     │
                            │  grava raw IMUTÁVEL      │
                            │  checkpoint a cada N     │
                            └────────────┬─────────────┘
                                         ▼
                            ┌──────────────────────────┐
                            │  ESTÁGIO 3: FILTROS §2   │
                            │  type==game              │
                            │  windows==True           │
                            │  coming_soon==False      │
                            │  release_date >= 2005    │
                            │  evidências 3D / 2D      │
                            │  conflitos -> PRESERVA   │
                            │  cada descarte -> ledger │
                            └────────────┬─────────────┘
                                         ▼
                            ┌──────────────────────────┐
                            │  ESTÁGIO 4: PARSE REQS   │
                            │  minimum_*  (separado)   │
                            │  recommended_* (separado)│
                            │  raw SEMPRE preservado   │
                            └────────────┬─────────────┘
                                         ▼
                            ┌──────────────────────────┐
                            │  ESTÁGIO 5: VALIDATE     │
                            │  ESTÁGIO 6: EXPORT       │
                            │  dataset.json tabular    │
                            └──────────────────────────┘
```

**Princípio de separação:** os estágios 3–6 operam **exclusivamente sobre
`data/raw/`**, nunca sobre a rede. Consequência prática: mudar um critério de
filtragem, corrigir um bug do parser ou ajustar o corte de 2005 **não exige
re-raspar a Steam**. Isso atende ao §12 (reprodutibilidade) e ao §5
(preservação do texto bruto) do projeto.

---

## 3. Contratos dos endpoints

### 3.1. Taxonomia (1 requisição por execução)

```
GET https://api.steampowered.com/IStoreService/GetTagList/v1/?language=english
-> {"response": {"tags": [{"tagid": 4191, "name": "3D"}, ...]}}  # 446 tags
```

Snapshot gravado em `data/raw/tag_taxonomy_<data>.json`. Se qualquer nome do
`filters.yaml` não resolver, o pipeline **aborta com erro** — nunca prossegue com
um filtro silenciosamente vazio (foi exatamente esse o erro que cometi na
sondagem inicial ao supor que 4166 era `3D`).

### 3.2. Discovery (1 query por tag positiva, paginada)

```
GET https://store.steampowered.com/search/results/
    ?query=&start=<N>&count=50&infinite=1&json=1
    &category1=998 &os=win &tags=<tagid> &untags=492 &sort_by=Released_DESC
```

Notas obrigatórias, medidas na FASE 1:
- `tags` múltiplas = **AND**. A união das positivas do §2.6 **precisa** ser feita
  no cliente, uma query por tag.
- `untags` = complemento exato (validado: 25.036 + 26.724 = 51.760).
- `count=50` é o máximo por página; `start` funciona até ≥45.000.
- extrair `data-ds-appid` e `data-ds-tagids` do `results_html`.
- **Rate limit severo:** 429 na 17ª req a 0,35 s. Alvo: **1 req / 4,0 s**.

### 3.3. Scrape — metadados e requisitos (1 req por candidato)

```
GET https://store.steampowered.com/api/appdetails?appids=<id>&l=english&cc=us
```

- `l=english` **obrigatório** (estabiliza os rótulos dos requisitos).
- **HTTP 200 com `success:false`** para app_id inválido — checar `success`,
  jamais o status code.
- `recommendations.total` = review_count oficial (Q4).

### 3.4. Scrape — tags (1 req por candidato, decisão Q3)

```
GET https://store.steampowered.com/app/<id>/
-> regex: InitAppTagModal\(\s*\d+\s*,\s*(\[.*?\])\s*,   ->  20 tags {tagid,name}
```

- Persistir **apenas as tags extraídas** (`<app_id>.tags.json`), não o HTML de
  270 KB — evita ~7 GB de lixo em disco para 27k jogos.
- Se redirecionar para `/login/?redir=agecheck/...` (caso HuniePop):
  `tags_source=AGE_GATE_LOGIN_REQUIRED`, tags `null`. **Não contornar.**
  Os requisitos do `appdetails` continuam sendo coletados.

### 3.5. Auditoria de viés (opcional, decisão Q1)

```
GET https://api.steampowered.com/IStoreService/GetAppList/v1/
    ?key=<STEAM_API_KEY>&include_games=true&max_results=50000&last_appid=<N>
```

- Key lida de `STEAM_API_KEY` (variável de ambiente, **nunca** commitada).
- Ausente a key: o estágio é **ignorado com WARNING**; o pipeline principal roda.
- Saída: `data/processed/bias_audit.json` — quantos app_ids do catálogo não
  aparecem no discovery via search, estratificado por ano de lançamento.

---

## 4. Filtros (`filters.yaml`, versionado)

```yaml
filter_version: "1.0.0"

app_type:
  allowed: [game]                    # §2.1
platform:
  require_windows: true              # §2.2
release:
  min_date: "2005-01-01"             # §2.3 (preliminar)
  exclude_coming_soon: true          # §2.4
  exclude_early_access: false        # §2.12

tags:
  positive_strong:                   # §2.6 — evidência forte de 3D
    [3D, First-Person, Third Person, Third-Person Shooter, FPS,
     3D Platformer, 3D Fighter, Immersive Sim, Walking Simulator,
     Automobile Sim, Flight, Space Sim, Looter Shooter, Hero Shooter,
     Arena Shooter, Rail Shooter]
  positive_secondary:                # §2.7 — complementar, nunca suficiente
    [Open World, Realistic, Cinematic, Action-Adventure, Action RPG,
     Souls-like, Survival Horror, Driving]
  negative_2d:                       # §2.8 — marca conflito, não exclui
    [2D, 2D Platformer, 2D Fighter]
  neutral_never_exclude: [2.5D]      # §2.9
  exclude_at_source: [Indie]         # §2.10 / Q2 -> untags=492
  negative_candidates:               # §2.11 — SÓ MARCA, não exclui
    [Pixel Graphics, Hand-drawn, Side Scroller, Text-Based, Visual Novel,
     Interactive Fiction, Point & Click, Card Game, Board Game, Hidden Object]
  never_exclude:                     # §2.12
    [Casual, Free to Play, Early Access]
  # §2.13: "Homemade" NÃO existe na taxonomia oficial -> regra não criada

reviews:
  metric: recommendations_total      # Q4
  min_threshold: null                # §2.14: NÃO aplicado ainda; só medir
```

### 4.1. Ordem de avaliação (§2.18)

`NOT_GAME` → `NO_WINDOWS_SUPPORT` → `UNRELEASED` → `BEFORE_2005` →
`NO_3D_EVIDENCE` → `TWO_DIMENSIONAL` → `BELOW_REVIEW_THRESHOLD`

O **primeiro** critério que reprovar define o `exclusion_reason`, garantindo
motivo único e determinístico por registro.

### 4.2. Resolução de conflitos (§2.8, §2.9, §2.16)

| Situação | Ação |
|---|---|
| tem positiva forte, sem `2D` | inclui, `inclusion_basis=STRONG_3D_TAG` |
| tem positiva forte **e** `2D` | **INCLUI**, `tag_conflict_3d_2d=true`, `needs_manual_review=true` (1.915 casos medidos) |
| tem `2.5D` | nunca exclui por isso (§2.9) |
| só positivas secundárias, sem forte | **INCLUI**, `inclusion_basis=SECONDARY_ONLY`, `needs_manual_review=true` — recall sobre precisão (§2.16) |
| negativa candidata (ex.: `Pixel Graphics`) | **apenas marca** em `negative_candidate_tags` |
| sem nenhuma positiva | exclui, `NO_3D_EVIDENCE` |

Cada registro carrega **`matched_positive_tags`** — as tags que causaram sua
inclusão, conforme exigido pelo §2.6 ("preservar quais tags produziram a
inclusão de cada registro").

### 4.3. Threshold de reviews (§2.14) — deliberadamente NÃO aplicado

`min_threshold: null`. O comando `python main.py analyze-thresholds` produzirá a
distribuição de reviews e a contagem de sobreviventes em 0 / 100 / 500 / 1.000 /
5.000, cumprindo os 4 passos que o §2.14 exige antes de fechar o valor. O corte
só será parametrizado após sua decisão.

---

## 5. Schema do registro (JSON tabular, §6)

```jsonc
{
  // --- identificação ---
  "app_id": 1091500,
  "name": "Cyberpunk 2077",
  "steam_url": "https://store.steampowered.com/app/1091500/",
  "type": "game",

  // --- lançamento (§2.3: data original SEMPRE preservada) ---
  "release_date_raw": "Dec 9, 2020",
  "release_date": "2020-12-09",
  "release_year": 2020,
  "coming_soon": false,

  // --- classificação ---
  "genres": ["RPG"],
  "categories": ["Single-player", "Steam Achievements"],
  "steam_tags": ["Cyberpunk", "Open World", "RPG", "FPS", "First-Person"],
  "steam_tag_ids": [4115, 1695, 122, 1663, 3839],
  "tags_source": "STORE_HTML",          // ou AGE_GATE_LOGIN_REQUIRED / UNAVAILABLE
  "developer": ["CD PROJEKT RED"],
  "publisher": ["CD PROJEKT RED"],

  // --- metadados auxiliares ---
  "review_count": 886203,               // recommendations.total (Q4)
  "review_count_metric": "recommendations_total",
  "metacritic_score": 86,
  "supported_languages_raw": "English<strong>*</strong>, French...",
  "supported_languages_count": 13,
  "is_free": false,
  "required_age": "17",
  "platform_windows": true,
  "platform_mac": true,
  "platform_linux": false,

  // --- REQUISITOS: prioridade máxima (§5) ---
  "pc_requirements": {
    "minimum": {
      "raw": "<strong>Minimum:</strong><br><ul class=\"bb_ul\">...",
      "os": "64-bit Windows 10", "os_legacy_flag": false,
      "cpu": "Core i7-6700 or Ryzen 5 1600",
      "ram": "12 GB RAM", "gpu": "GeForce GTX 1060 6GB",
      "directx": "Version 12", "storage": "70 GB available space",
      "network": null, "sound_card": null, "additional_notes": null,
      "unparsed_labels": []
    },
    "recommended": {
      "raw": "<strong>Recommended:</strong><br><ul class=\"bb_ul\">...",
      "os": "64-bit Windows 10", "os_legacy_flag": false,
      "cpu": "Core i7-12700 or Ryzen 7 7800X3D",
      "ram": "16 GB RAM", "gpu": "GeForce RTX 2060 SUPER or Radeon RX 5700 XT",
      "directx": "Version 12", "storage": "70 GB available space",
      "network": null, "sound_card": null, "additional_notes": null,
      "unparsed_labels": []
    }
  },
  "has_minimum_requirements": true,
  "has_recommended_requirements": true,   // §8: false NÃO é falha
  "requirements_markup_format": "A",      // A | B | UNKNOWN (§3.1 da FASE 1)

  // --- rastreabilidade do filtro (§2.17) ---
  "included_initially": true,
  "exclusion_reason": null,
  "inclusion_basis": "STRONG_3D_TAG",
  "matched_positive_tags": ["FPS", "First-Person", "Open World"],
  "is_indie": false,
  "has_2d_tags": false,
  "tag_conflict_3d_2d": false,
  "negative_candidate_tags": [],
  "needs_manual_review": false,

  // --- reprodutibilidade (§12) ---
  "filter_version": "1.0.0",
  "scraper_version": "0.1.0",
  "collected_at": "2026-08-30T14:23:11Z",
  "source": "steam_store_api",
  "source_urls": {
    "appdetails": "https://store.steampowered.com/api/appdetails?appids=1091500&l=english&cc=us",
    "store_page": "https://store.steampowered.com/app/1091500/"
  },
  "collection_status": "COMPLETE"       // COMPLETE | PARTIAL_NO_TAGS | FAILED
}
```

**Homogeneidade garantida:** todo registro tem **todas** as chaves; ausências são
`null` (§5: "não force valores ausentes"). Campos `minimum_*` e `recommended_*`
são estruturalmente isolados — **nenhum código copia mínimo para recomendado**.
O `export` oferece `--flat` para achatar `pc_requirements` em
`minimum_cpu`/`recommended_cpu` etc., conforme sugerido no §5 do projeto.

---

## 6. Parser de requisitos

Dado o §3.1 da FASE 1 (dois formatos de markup incompatíveis), a estratégia é:

1. **Detectar o formato.** Se `<li>` contém `</strong>` seguido de texto não
   vazio → **Formato A**. Se o `<strong>` engloba `rótulo: valor` → **Formato B**.
   Registrar em `requirements_markup_format` (`A`/`B`/`UNKNOWN`).
2. **Dividir por `<li>`** e, em cada item, separar rótulo e valor no primeiro
   `:`.
3. **Normalizar o rótulo** por mapa de sinônimos:
   - `os` ← `OS`, `OS *`, `Operating System`
   - `cpu` ← `Processor`, `CPU`
   - `ram` ← `Memory`, `RAM`
   - `gpu` ← `Graphics`, `Video Card`, `Graphics Card`, `Video`
   - `directx` ← `DirectX`, `DirectX Version`
   - `storage` ← `Storage`, `Hard Disk Space`, `Hard Drive`, `HDD`, `Available Space`
   - `network` ← `Network`, `Internet Connection`
   - `sound_card` ← `Sound Card`, `Sound`
   - `additional_notes` ← `Additional Notes`
4. **`OS *` → `os_legacy_flag=true`** e rótulo normalizado para `os` (o asterisco
   é sinal, não ruído).
5. **Rótulo desconhecido → `unparsed_labels`**, nunca descartado silenciosamente.
   Essa lista é o instrumento de auditoria do parser.
6. **Preservar o `raw`** sempre e integralmente.

**Escopo deliberadamente limitado (§5):** o parser **não** converte "16 GB RAM"
em `16`, nem mapeia GPUs para benchmarks. Extrai o texto do campo. A
normalização de hardware pertence ao Data Wrangling; nesta etapa a prioridade é
fidelidade, não interpretação.

---

## 7. Checkpoint e retomada (§7, §11)

`data/checkpoints/scrape_state.json`, uma entrada por app_id:

```jsonc
{
  "meta": {"filter_version": "1.0.0", "scraper_version": "0.1.0",
           "started_at": "...", "last_updated": "..."},
  "apps": {
    "1091500": {"status": "COMPLETE", "attempts": 1,
                "last_error": null, "last_attempt_at": "...",
                "has_appdetails": true, "has_tags": true},
    "999999999": {"status": "FAILED", "attempts": 3,
                  "last_error": "APPDETAILS_SUCCESS_FALSE", ...},
    "339800": {"status": "PARTIAL_NO_TAGS", "attempts": 1,
               "last_error": "AGE_GATE_LOGIN_REQUIRED", ...}
  }
}
```

Estados: `PENDING`, `IN_PROGRESS`, `COMPLETE`, `PARTIAL_NO_TAGS`, `FAILED`,
`EXCLUDED_BY_FILTER`.

Garantias:
- **Escrita atômica:** grava em `.tmp` + `os.replace()`. Interrupção durante a
  escrita nunca corrompe o checkpoint.
- **Idempotência:** app com `raw/appdetails/<id>.json` presente e checkpoint
  `COMPLETE` é **pulado sem requisição de rede**. Reexecutar é seguro e barato.
- **Flush periódico:** a cada `checkpoint_every: 25` apps, e no `SIGINT`.
- **`Ctrl+C` limpo:** handler de `KeyboardInterrupt` faz flush do checkpoint e
  encerra com resumo. Retomar no jogo 4.731 continua no 4.732.

---

## 8. Erros e resiliência (§10)

| Erro | Classificação | Ação |
|---|---|---|
| `Timeout`, `ConnectionError` | transitório | retry, backoff exponencial + jitter |
| HTTP **429** | transitório | backoff **cego** (sem `Retry-After`): 4→8→16→32→64 s |
| HTTP **5xx** | transitório | retry até `max_retries` (3) |
| HTTP **404** | permanente | `FAILED:HTTP_404`, sem retry |
| HTTP **403** | permanente | `FAILED:HTTP_403` (ex.: API key inválida) |
| Redirect p/ `agecheck` | permanente-parcial | `PARTIAL_NO_TAGS`, requisitos preservados |
| JSON inválido | transitório-1x | 1 retry, depois `FAILED:INVALID_JSON` |
| `success:false` | **dado ausente** | `FAILED:APPDETAILS_SUCCESS_FALSE`, sem retry |
| `InitAppTagModal` ausente | falha de parsing | `PARTIAL_NO_TAGS` + WARNING (sinal de mudança de estrutura) |
| Rótulo desconhecido | falha de parsing | `unparsed_labels`, registro **válido** |
| Sem `recommended` | **dado ausente na Steam** | `has_recommended_requirements=false`, **não é erro** |
| app_id duplicado | dado | deduplicação por app_id, WARNING |
| Falha de escrita | fatal local | escrita atômica; aborta com erro claro |
| `KeyboardInterrupt` | controlado | flush + resumo + exit 0 |

**Distinção exigida pelo §8**, materializada em campos separados:
`collection_status` (ausência real vs. página inacessível vs. falha HTTP vs.
falha de parsing) e `exclusion_reason` (exclusão metodológica). **Nunca
confundidos.**

Uma falha individual **nunca** interrompe o lote: cada app é processado em
`try/except` isolado, o erro vai para o checkpoint e o loop continue (§10).

### 8.1. Rate limiting (calibrado pelas medições da FASE 1)

```yaml
rate_limits:
  search:      {min_interval_s: 4.0,  max_retries: 5, backoff_base_s: 4.0}
  appdetails:  {min_interval_s: 0.6,  max_retries: 3, backoff_base_s: 2.0}
  store_html:  {min_interval_s: 1.0,  max_retries: 3, backoff_base_s: 2.0}
timeouts: {connect_s: 10, read_s: 30}
```

Justificativa: 429 na 17ª req do search a 0,35 s → 4,0 s dá ~11× de margem.
`appdetails` aguentou 30/30 a 0,35 s → 0,6 s é conservador. Sessão HTTP única e
reutilizada (keep-alive), User-Agent identificando a pesquisa acadêmica.

---

## 9. CLI (§11)

```bash
python main.py tags                       # baixa/valida a taxonomia; resolve filters.yaml
python main.py discover                   # estágio 1 -> candidates.json (+ ledger Indie)
python main.py discover --resume
python main.py scrape                     # estágio 2, retomável
python main.py scrape --limit 20          # FASE 3: protótipo
python main.py scrape --start-from 1091500
python main.py scrape --app-ids 1091500,1245620,220,730,105600
python main.py resume                     # alias de scrape (pula COMPLETE)
python main.py retry-failed               # só status FAILED
python main.py filter                     # estágio 3, offline sobre raw/
python main.py validate                   # estágio 5 (§8)
python main.py export [--flat]            # estágio 6 -> dataset.json
python main.py analyze-thresholds         # §2.14: distribuição de reviews
python main.py audit-bias                 # Q1: exige STEAM_API_KEY
python main.py status                     # resumo do checkpoint
```

---

## 10. Logging (§9)

Console verbose + `data/logs/run_<timestamp>.log` (`DEBUG` em arquivo, `INFO` em
console). Formato dos progressos:

```
[TAGS]      446 tags oficiais carregadas | 43 tags do filtro resolvidas | 1 inexistente: Homemade
[TAGS]      AVISO: Rail Shooter (3954) existe mas retorna 0 jogos
[DISCOVERY] tag 3D (4191) untags=Indie(492) -> total_count=26,724
[DISCOVERY] tag 3D          | pág 012/535 | 600 app_ids | 2.2% | decorrido 00:00:48
[DISCOVERY] união de 16 tags positivas -> 41,203 app_ids únicos
[DISCOVERY] ledger de excluídos por Indie -> 25,036 app_ids
[SCRAPE 00431/41203 | 1.0%] app_id=1091500 | Cyberpunk 2077
   [OK] appdetails 35.8 KB | type=game | release=2020-12-09 | reviews=886,203
   [OK] 20 tags via STORE_HTML
   [OK] requisitos RECOMENDADOS capturados (formato A)
[SCRAPE 00432/41203 | 1.0%] app_id=220 | Half-Life 2
   [SKIP] BEFORE_2005 (release_date_raw="Nov 16, 2004")
[SCRAPE 00433/41203 | 1.0%] app_id=339800 | HuniePop
   [WARN] AGE_GATE_LOGIN_REQUIRED -> tags indisponíveis; requisitos preservados
   [PARTIAL] collection_status=PARTIAL_NO_TAGS
[RATE] HTTP 429 no search -> backoff 8.0s (tentativa 2/5)
[CHECKPOINT] 431 completos | 12 parciais | 5 falhas | 3 excluídos | 40,752 restantes
             decorrido 00:07:12 | ~0.98s/app | ETA 11h05m
[FILTER] entrada 41,203
[FILTER]   NOT_GAME              ->  -1,204  (40,000 restantes)
[FILTER]   BEFORE_2005           ->  -2,118  (37,882 restantes)
[FILTER]   NO_3D_EVIDENCE        ->  -4,003  (33,879 restantes)
[FILTER]   conflitos 3D+2D preservados: 1,915 (needs_manual_review=true)
[EXPORT] 33,879 registros -> data/processed/dataset.json (UTF-8, 412 MB)
```

Nunca imprime HTML/JSON íntegro — apenas tamanhos, contagens e trechos curtos.

---

## 11. Testes (§13)

| Arquivo | Cobre |
|---|---|
| `test_requirements_parser.py` | Formato A (Cyberpunk); **Formato B (Terraria)**; `recommended` ausente (CS2); `OS *` → `os_legacy_flag`; sinônimos `Hard Disk Space`/`Video Card`; rótulo desconhecido → `unparsed_labels`; RAM e armazenamento; **mínimo nunca vaza para recomendado** |
| `test_filters.py` | `type=dlc` → `NOT_GAME`; sem Windows; `coming_soon`; 2004 → `BEFORE_2005`; **conflito 3D+2D preservado**; `2.5D` nunca exclui; secundária isolada inclui com `needs_manual_review`; negativa candidata só marca; `Casual`/`F2P`/`Early Access` nunca excluem; ordem de `exclusion_reason` |
| `test_tags.py` | resolução nome→tagid; **falha ruidosa em tag inexistente** (`Homemade`); extração do `InitAppTagModal` |
| `test_storage.py` | escrita atômica; checkpoint/resume; idempotência; JSON serializável; UTF-8 |
| `test_validators.py` | §8 completo |
| `test_steam_client.py` | 429 → backoff (mock); 5xx → retry; 404 → sem retry; `success:false` |

Fixtures pequenas, extraídas dos payloads reais de `investigation/samples/`.

---

## 12. Critérios de aceite (§16) — mapeamento

| # | Critério | Onde é atendido |
|---|---|---|
| 1 | população descoberta | `discovery.py` |
| 2 | filtros parametrizados | `config/filters.yaml` + `filter_version` |
| 3 | não-pertinentes removidos antes do scraping | `category1=998&os=win&untags=492` no servidor |
| 4 | recomendados separados dos mínimos | schema §5, testado |
| 5 | textos brutos preservados | `raw` + `data/raw/` imutável |
| 6 | falha individual não derruba o pipeline | §8 |
| 7 | retries e backoff | `steam_client.py` |
| 8 | checkpoints e resume | §7 |
| 9 | reprocessamento de falhas | `retry-failed` |
| 10 | interromper e retomar | escrita atômica + `SIGINT` |
| 11 | logs persistentes | `data/logs/` |
| 12 | validações | `validators.py` |
| 13 | testes | `tests/` |
| 14 | JSON tabular | `export.py --flat` |
| 15 | metodologia documentada | `docs/` |
| 16 | amostra confrontada com a Steam | **FASE 4** |

---

## 13. Próximos passos

- **FASE 3 (protótipo):** implementar o pipeline e rodar `scrape --app-ids` sobre
  uma amostra heterogênea de ~15 jogos (2005–2024; Formato A e B; sem
  recomendados; DLC; F2P; age gate; conflito 3D+2D).
- **FASE 4 (validação):** conferir manualmente os requisitos **recomendados**
  contra as páginas reais da Steam.
- **FASE 5:** só após a validação, `discover` completo e `scrape` em escala.
