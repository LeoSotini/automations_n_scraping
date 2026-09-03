# Registro de Decisões Metodológicas

Conforme o §12 do projeto ("não altere silenciosamente critérios metodológicos
durante a execução"), toda decisão que afete a população do estudo é registrada
aqui, com data, alternativas consideradas e evidência empírica que a embasou.

---

## D1 — Estratégia de discovery: híbrida
**Data:** 2026-08-30 · **Decidido por:** Leonardo (Q1) · **Status:** ativa

`ISteamApps/GetAppList/v2` está descontinuado (HTTP 404 em 5 variantes testadas).
Discovery principal via `store.steampowered.com/search/results/`, com auditoria
opcional do viés via `IStoreService/GetAppList/v1` (exige `STEAM_API_KEY`).

**Alternativas:** só search (mais rápido, viés não quantificado); só API key
(~35× mais requisições).
**Limitação aceita:** o search só enumera jogos atualmente visíveis na loja.
Jogos deslistados desaparecem, e a probabilidade de deslistagem cresce com a
idade → viés de sobrevivência **correlacionado com o tempo**, que é o eixo do
estudo. A auditoria quantificará a magnitude por ano de lançamento.
**A declarar na metodologia do TCC.**

---

## D2 — `Indie` excluída na fonte
**Data:** 2026-08-30 · **Decidido por:** Leonardo (Q2) · **Status:** ativa

`untags=492` no discovery. Validado empiricamente como complemento exato
(25.036 + 26.724 = 51.760, diferença zero).

**Impacto medido:** remove **25.036 jogos, 48,4% de toda a população 3D**.
**Alternativa rejeitada:** marcar `is_indie` e coletar tudo (custo ~2×).
**Tensão declarada:** conflita parcialmente com o §2.16 (priorizar recall). A tag
`Indie` é auto-declarada e não mede orçamento — títulos como Hades, Hollow Knight
e Valheim são excluídos por ela.
**Mitigação implementada:** `data/processed/ledger_excluded_indie.json` grava os
app_ids excluídos com `exclusion_reason=INDIE` e `filter_version`. Revisar o
§2.10 no futuro exige apenas o estágio de scraping sobre o ledger — **sem repetir
o discovery**. A decisão permanece reversível e auditável (§2.17).

---

## D3 — Tags a partir do HTML da página do jogo
**Data:** 2026-08-30 · **Decidido por:** Leonardo (Q3) · **Status:** ativa

20 tags por jogo via `InitAppTagModal` no HTML. O `appdetails` **não** retorna
tags (verificado em 9 jogos); `genres`/`categories` não são tags.

**Alternativa rejeitada:** os 7–8 `data-ds-tagids` grátis do search — confirmado
como **subconjunto** das 20, produziria falsos negativos, proibidos pelo §2.16.
**Custo:** +1 requisição por candidato (~3,7 h → ~7,4 h para 27k jogos).
**Limitação:** jogos com age gate por login (ex.: HuniePop 339800) ficam sem
tags → `tags_source=AGE_GATE_LOGIN_REQUIRED`. Não será contornado.

---

## D4 — `recommendations.total` como métrica oficial de reviews
**Data:** 2026-08-30 · **Decidido por:** Leonardo (Q4) · **Status:** ativa

Vem grátis no `appdetails`; nenhuma chamada a `appreviews` será feita.

**Divergência documentada:** para CS2, `recommendations.total`=5.245.985 vs.
`appreviews.total_reviews`=9.826.091. São métricas distintas; o escopo exato de
`recommendations` não é documentado pela Steam. O threshold do §2.14 deve ser
interpretado **sempre** nesta escala, nunca comparado a números publicados que
usem a outra métrica.
**A declarar na metodologia do TCC.**

---

## D5 — Regra `NOT Homemade` não criada
**Data:** 2026-08-30 · **Base:** evidência empírica · **Status:** ativa

O §2.13 exigia confirmar a existência da tag antes de criar a regra.
**Confirmado: "Homemade" não existe** entre as 446 tags oficiais
(`IStoreService/GetTagList/v1`). A regra não foi criada.

---

## D6 — Threshold de reviews não aplicado nesta versão
**Data:** 2026-08-30 · **Base:** §2.14 · **Status:** pendente de decisão

`min_threshold: null`. O §2.14 exige 4 passos antes de fechar o valor. O comando
`analyze-thresholds` produzirá a distribuição e as contagens de sobreviventes
(0/100/500/1.000/5.000) para decisão posterior e justificada.

---

## D7 — Tags negativas candidatas apenas marcadas, nunca excluídas
**Data:** 2026-08-30 · **Base:** §2.11 + evidência empírica · **Status:** ativa

Medido: **2.389 jogos são simultaneamente `3D` e `Pixel Graphics`**, confirmando
o alerta do §2.11. Nenhuma das 10 negativas candidatas é exclusão rígida; todas
vão para `negative_candidate_tags`, permitindo quantificar o impacto de cada
regra antes de efetivá-la.

---

## D8 — Conflitos 3D+2D preservados
**Data:** 2026-08-30 · **Base:** §2.8, §2.16 + evidência empírica · **Status:** ativa

Medido: **1.915 jogos têm simultaneamente `3D` e `2D`**. Não são descartados;
recebem `tag_conflict_3d_2d=true` e `needs_manual_review=true` para resolução
posterior, conforme o §2.8 exige.

---

## D9 — Age gate: duas variantes, tratamento distinto
**Data:** 2026-08-31 · **Base:** evidência empírica da FASE 3 · **Status:** ativa

Descoberto durante o protótipo: **7 de 17 páginas** retornaram ~51 KB sem
`InitAppTagModal`. O diagnóstico (`investigation/diag_missing_tags.py`) revelou
que a Steam usa **duas** variantes de bloqueio por idade, não uma:

| Variante | Comportamento medido | Exige autenticação? |
|---|---|---|
| **(a) interstitial na página** | HTTP 200, ~51 KB, marcador `agecheck` no corpo, **sem redirect** | **não** — auto-declaração de data de nascimento |
| **(b) redirect para login** | `/login/?redir=agecheck/app/<id>/` | **sim** |

A implementação inicial só reconhecia (b), e por isso rotulou (a) erradamente
como "possível mudança de estrutura da página".

**Decisão:**
* variante **(b) nunca é contornada** (item 3 do projeto). Registra
  `tags_source=AGE_GATE_LOGIN_REQUIRED`; os requisitos continuam sendo coletados
  via `appdetails`, que não é bloqueado.
* variante **(a)** é controlada por `accept_age_gate_interstitial` em
  `settings.yaml` (atualmente `true`), com `tags_source` registrando
  `STORE_HTML_AFTER_AGE_GATE` para plena auditabilidade.

**Fundamentos da decisão sobre (a):**
1. não há CAPTCHA, autenticação nem mecanismo anti-bot — apenas auto-declaração;
2. o `robots.txt` da Steam **permite** `/app/` (bloqueia somente `/share/`,
   `/news/externalpost/`, URLs com token, `/login/?guestpasskey`,
   `/join/?redir`, `/account/ackgift/`, `/email/`, `/widget/`);
3. os mesmos dados de requisitos são servidos **sem qualquer gate** pela própria
   API pública da Steam (`appdetails`) — o gate incide apenas no HTML;
4. **argumento metodológico decisivo:** o gate incide sobre títulos mature, que
   se concentram em jogos de ação AA/AAA — exatamente a população-alvo do
   estudo. Omitir suas tags introduziria viés sistemático **contra** o objeto da
   pesquisa. Medido no protótipo: 6 de 10 jogos com requisitos recomendados
   (Cyberpunk 2077, ELDEN RING, The Witcher 3, Hades, Stardew Valley e o DLC).

**Reversível:** definir `accept_age_gate_interstitial: false` faz o pipeline
registrar `tags_source=AGE_GATE_INTERSTITIAL` e omitir as tags, sem qualquer
outra alteração de código.

---

## D10 — Jogos 2D sem tag forte de 3D são excluídos
**Data:** 2026-08-31 · **Decidido por:** Leonardo · **Status:** ativa
**Efeito:** `filter_version` 1.0.0 → **1.1.0**

Descoberto na validação da FASE 4. A regra anterior, derivada dos §2.7 e §2.16,
incluía jogos com **apenas** tags positivas secundárias
(`inclusion_basis=SECONDARY_ONLY`), o que admitia jogos inequivocamente
bidimensionais:

| Jogo | Tags relevantes | Resultado anterior |
|---|---|---|
| Terraria | `2D`, `Pixel Graphics`, `Open World` | incluído indevidamente |
| Hollow Knight | `2D`, `Metroidvania`, `Souls-like` | incluído indevidamente |
| Stardew Valley | `Pixel Graphics`, `2D`, `RPG` | incluído indevidamente |

Magnitude medida no universo Steam (Windows, jogos, `NOT Indie`):

| Interseção | Jogos | | Interseção | Jogos |
|---|---:|---|---|---:|
| `Action-Adventure` ∩ `2D` | 3.820 | | `Realistic` ∩ `2D` | 502 |
| `Action RPG` ∩ `2D` | 2.025 | | `Cinematic` ∩ `2D` | 477 |
| `Open World` ∩ `2D` | 1.386 | | `Survival Horror` ∩ `2D` | 415 |
| `Souls-like` ∩ `2D` | 539 | | `Driving` ∩ `2D` | 147 |

**Regra adotada:**

```
tag forte de 3D + 2D        -> INCLUI  (conflito preservado, §2.8)   [1.915 casos]
secundária + 2D             -> EXCLUI  (TWO_DIMENSIONAL)             <- D10
secundária, sem 2D          -> INCLUI  (needs_manual_review=true)
tag forte de 3D, sem 2D     -> INCLUI
```

**Fundamento:** sem nenhuma tag forte de 3D, a evidência negativa do §2.8
prevalece sobre a evidência complementar do §2.7 — e o próprio §2.7 afirma que
as secundárias "NÃO garantem individualmente que o jogo seja 3D". O §2.8
continua integralmente respeitado nos casos que ele descreve: conflitos com
evidência **forte** de 3D seguem preservados com `tag_conflict_3d_2d=true`.

**Limitação aceita:** um jogo tridimensional cuja tag `3D` (ou outra forte) não
esteja entre as 20 tags principais, mas que tenha `2D`, será excluído. O
`ledger_exclusions.json` registra todos com `exclusion_reason=TWO_DIMENSIONAL`,
permitindo auditoria e reversão sem re-raspar a Steam.

**Verificado no protótipo:** removeu exatamente Terraria, Hollow Knight e
Stardew Valley. Hades permaneceu (`SECONDARY_ONLY`) porque `2D` não está entre
suas 20 tags — e Hades de fato usa renderização 3D isométrica, o que torna a
inclusão correta.

---

## D11 — `sort_by` REMOVIDO do discovery (bug de seleção)
**Data:** 2026-08-31 · **Base:** evidência empírica · **Status:** ativa

Descoberto ao rodar o discovery em modo piloto: `3D NOT Indie` retornou 15.744,
mas a FASE 1 havia medido 26.755 para o mesmo recorte. A única diferença era o
parâmetro `sort_by=Released_DESC`, que eu havia incluído em `settings.yaml`
presumindo que ordenar fosse inócuo.

Medição (`investigation/diag_sort_by.py`):

| Query | `total_count` |
|---|---:|
| `3D`, sem `sort_by` | 51.841 |
| `3D`, `sort_by=Relevance` | 51.841 |
| `3D`, `sort_by=Released_DESC` | 31.273 (−39,7%) |
| `3D`, `sort_by=Reviews_DESC` | 16.670 (−67,8%) |
| `3D NOT Indie`, sem `sort_by` | **26.755** |
| `3D NOT Indie`, `sort_by=Released_DESC` | **15.744 (−41,2%)** |
| `3D NOT Indie`, `sort_by=Reviews_DESC` | 9.012 (−66,3%) |

**O `sort_by` não apenas ordena: ele FILTRA** itens que não possuem a chave de
ordenação. Perderia **41% da população**, concentrando a perda justamente em
entradas com data de lançamento ausente ou irregular — num estudo cujo eixo é a
data. Seria um viés de seleção silencioso e devastador.

**Decisão:** `sort_by: ""` em `settings.yaml`; o parâmetro só é enviado se
explicitamente configurado. `Relevance` seria seguro (população idêntica), mas
não traz benefício, já que a enumeração percorre todas as páginas.

**Mitigação do risco decorrente:** sem ordenação explícita, a ordem entre
páginas pode ser instável, o que arriscaria nunca ver alguns itens. O discovery
passou a medir e registrar **cobertura por tag** (`distinct_seen / total_count`)
e emite WARNING abaixo de 95%. Como a união é cumulativa e deduplicada por
app_id, reexecutar `discover` recupera itens faltantes.

---

## D12 — Ritmo do scrape: 2 s entre jogos
**Data:** 2026-08-31 · **Decidido por:** Leonardo, após piloto · **Status:** ativa

O ritmo foi decidido com base em piloto empírico, não em estimativa.

**Piloto executado:** 300 jogos reais (`scrape --from-pilot --limit 300`) com
`sleep_between_apps_s=2.0`:

| Métrica | Resultado |
|---|---|
| Jogos processados | 293 completos, 0 parciais, 0 falhas, 7 pulados |
| Requisições HTTP | **700** |
| Retries | **0** |
| HTTP 429 | **0** |
| Ritmo efetivo | 3,22 s/app |
| Duração | 15 min 40 s |

**Custo projetado por ritmo** (cenário intermediário, ~63.500 jogos):

| Sleep | Duração | req/s |
|---:|---:|---:|
| 0 s | 28 h (1,2 d) | 1,25 |
| 1 s | 46 h (1,9 d) | 0,77 |
| **2 s** | **64 h (2,6 d)** | **0,56** |
| 3 s | 81 h (3,4 d) | 0,43 |
| 8 s | 169 h (7,1 d) | 0,21 |

Faixa completa a 2 s: **44 h (otimista) a 89 h (pessimista)**.

**Fundamento:** 700 requisições sustentadas sem um único 429 a 0,56 req/s é
evidência direta de que 2 s é seguro para `/api/appdetails` e `/app/<id>/`. O
endpoint sensível é o `/search/results/` (429 na 17ª requisição na FASE 1),
usado apenas no discovery e configurado a 4,0 s/req.

**Rede de segurança:** se a Steam apertar o limite, o backoff exponencial cego
(2→4→8 s, teto 120 s) degrada o ritmo em vez de falhar, e o 429 fica registrado
no log e no contador `rate_limited`. Aumentar o valor em `settings.yaml` não
exige alteração de código, e a coleta é retomável.

---

## D13 — Sinônimo `Supported OS` e rótulos deliberadamente não mapeados
**Data:** 2026-08-31 · **Base:** auditoria do piloto de 310 jogos · **Status:** ativa

A auditoria de `unparsed_labels` (`investigation/audit_unparsed_labels.py`)
revelou 10 rótulos distintos não mapeados. Análise:

| Rótulo | Ocorr. | Ação |
|---|---:|---|
| `VR Support` | 17 | **não mapear** — fora do escopo dos itens 4/5 |
| `Supported OS` | 4 | **corrigido** → `os` (Left 4 Dead, Resident Evil 5) |
| `Video Card Memory` | 2 | **não mapear** — é VRAM, não o modelo da GPU |
| `Display`, `Peripherals`, `Input` | 5 | **não mapear** — fora do escopo |
| `Video Card (duplicado)` | 2 | guarda de rótulo repetido funcionando |
| `Minimum`, `Recommended` | 2 | prosa em páginas antigas (app 10) |
| `Further information` | 1 | fora do escopo |

Todos permanecem visíveis em `unparsed_labels` — nenhum é descartado.

**Cobertura verificada em escala** (289 jogos com requisitos recomendados):
`cpu`/`ram`/`gpu` 94,8%, `os` 93,8%, `storage` 91,7%, `directx` 75,4%.
As lacunas foram investigadas caso a caso
(`investigation/diag_field_gaps.py`) e são **ausência real na Steam**, não falha
de parser. Dois padrões relevantes para o Data Wrangling:

1. alguns jogos publicam como "recomendado" apenas a linha
   `"Requires a 64-bit processor and operating system"`, sem nenhum campo
   (VRChat, Call of Duty: WWII, Half-Life: Alyx);
2. **o armazenamento às vezes aparece em `Additional Notes`** em vez do rótulo
   `Storage` — ex.: Call of Duty (1938090),
   `"SSD with 161 GB available space at launch"`. O dado está preservado em
   `additional_notes` e deve ser recuperado no Data Wrangling.
