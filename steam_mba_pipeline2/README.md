# steam-req-pipeline

Pipeline reprodutível de coleta de **requisitos de hardware de jogos de PC** na
Steam, para TCC de MBA em Data Science & Analytics.

Objeto da pesquisa: a evolução dos requisitos computacionais de jogos de PC ao
longo do tempo, com foco na hipótese de crescimento desproporcional de exigências
de hardware e armazenamento em jogos modernos.

**O dado central são os REQUISITOS RECOMENDADOS.** Os requisitos mínimos são
coletados e preservados, mas mantidos rigorosamente separados. Nenhuma etapa do
código copia, mistura ou infere um a partir do outro.

---

## Instalação

```powershell
python -m pip install -r requirements.txt      # ou: python -m pip install -e .
python -m pip install pytest                   # para os testes
```

Dependências mínimas: `requests` e `PyYAML`. O parsing de HTML usa apenas a
biblioteca padrão (regex + `html.unescape`) — as estruturas-alvo
(`InitAppTagModal`, itens `<li>` dos requisitos) são simples e regulares, e
evitar `bs4`/`lxml` reduz a superfície de dependências.

---

## Uso

```powershell
python main.py tags                 # baixa a taxonomia oficial e resolve filters.yaml
python main.py discover             # estágio 1: enumera candidatos
python main.py scrape --prototype   # estágio 2 sobre a amostra heterogênea (17 jogos)
python main.py scrape               # estágio 2 sobre todos os candidatos
python main.py resume               # retoma de onde parou
python main.py retry-failed         # reprocessa apenas as falhas
python main.py filter               # estágio 3: aplica a metodologia (offline)
python main.py validate             # estágio 5: valida os registros
python main.py export --flat        # estágio 6: dataset.json tabular
python main.py analyze-thresholds   # item 2.14: distribuição de reviews
python main.py audit-bias           # decisão D1 (requer STEAM_API_KEY)
python main.py status               # resumo do checkpoint
python -m pytest -q                 # 154 testes
```

Limitar testes: `--limit N`, `--start-from APP_ID`, `--app-ids 730,620,220`.

---

## Arquitetura

O pipeline é dividido em dois domínios com uma fronteira rígida:

```
REDE                                   |  OFFLINE (sobre data/raw)
---------------------------------------|--------------------------------------
tags      taxonomia oficial            |  filter    metodologia do item 2
discover  /search/results/  ─┐         |  parse     requisitos A/B
scrape    appdetails + HTML ─┴─> raw ──┼─>  validate  item 8
                                       |  export    dataset.json tabular
```

**Consequência de projeto:** os estágios offline nunca tocam a rede. Alterar um
critério de filtragem, corrigir um bug do parser ou mover o corte de 2005
**não custa uma única requisição à Steam**. É isso que torna operacionais a
reprodutibilidade (item 12) e a preservação do texto bruto (item 5).

```
config/    settings.yaml (operação)  |  filters.yaml (METODOLOGIA, versionada)
data/raw/  payloads brutos, imutáveis — a fonte de verdade
data/processed/  candidates, ledgers, dataset, relatórios
data/checkpoints/  estado retomável por app_id
data/logs/  log persistente por execução
src/steamreq/  módulos do pipeline
investigation/  scripts da FASE 1 e validações (reprodutibilidade)
docs/  investigação, arquitetura e decisões metodológicas
```

---

## Metodologia de filtragem

A metodologia **é** o `config/filters.yaml`. Alterá-lo exige incrementar
`filter_version`, que é gravado em cada registro do dataset.

Ordem de avaliação (item 2.18) — o **primeiro** critério reprovado define o
`exclusion_reason`, garantindo motivo único e determinístico:

```
type == "game"          -> NOT_GAME
Windows == True         -> NO_WINDOWS_SUPPORT
coming_soon == False    -> UNRELEASED
release_date >= 2005    -> BEFORE_2005
evidência de 3D         -> NO_3D_EVIDENCE
tags 2D                 -> TWO_DIMENSIONAL
threshold de reviews    -> BELOW_REVIEW_THRESHOLD   (não aplicado nesta versão)
```

Princípios que o código respeita explicitamente:

| Item | Regra | Implementação |
|---|---|---|
| 2.16 | priorizar recall | em dúvida, **inclui** e marca `needs_manual_review` |
| 2.17 | nenhum descarte sem motivo | `ledger_exclusions.json` + `exclusion_reason` |
| 2.6 | preservar a causa da inclusão | `matched_positive_tags` |
| 2.8 | conflito 3D+2D não exclui | `tag_conflict_3d_2d=true`, preservado |
| 2.9 | `2.5D` nunca exclui | validado no carregamento da config |
| 2.11 | negativas candidatas só marcam | `negative_candidate_tags` |
| 2.12 | `Casual`/`F2P`/`Early Access` nunca excluem | rejeitado na config |
| 2.13 | `Homemade` | **não existe** na taxonomia; regra não criada |
| 2.14 | threshold de reviews | `null`; use `analyze-thresholds` |

`Indie` é excluída **na fonte** (`untags=492`, decisão D2), com os app_ids
excluídos gravados em `ledger_excluded_indie.json` — o que mantém a decisão
**reversível sem repetir o discovery**.

---

## Fontes de dados (verificadas empiricamente em 2026-08-30)

| Endpoint | Uso | Auth |
|---|---|---|
| `IStoreService/GetTagList/v1` | 446 tags oficiais | não |
| `store/search/results/` | discovery filtrado | não |
| `store/api/appdetails` | metadados + **requisitos** + review_count | não |
| `store/app/<id>/` | as 20 tags (`InitAppTagModal`) | não |
| `IStoreService/GetAppList/v1` | auditoria de viés (opcional) | **sim** |

Armadilhas confirmadas e tratadas no código:

- **`ISteamApps/GetAppList/v2` está morto (HTTP 404).** Cinco variantes testadas.
- **`appdetails` não retorna tags.** `genres`/`categories` não são tags.
- **`tags` múltiplas no search é AND, não OR** → a união das tags positivas é
  feita no cliente, uma query por tag.
- **`untags` é complemento exato** (validado: 25.036 + 26.724 = 51.760).
- **app_id inválido responde HTTP 200** com `success:false`. O status code é
  inútil como sinal; o código checa `success`.
- **DLC tem `pc_requirements` idênticos ao jogo-base** → o filtro `type=="game"`
  é indispensável.
- **`pc_requirements` tem ao menos dois formatos de markup incompatíveis**
  (ver abaixo).
- **Rate limit:** `/search/` deu 429 na 17ª requisição a 0,35 s;
  `/api/appdetails` aguentou 30/30. Limites são **por endpoint**. Não há header
  `Retry-After` — o backoff é cego, exponencial com jitter.

---

## Parsing dos requisitos

Dois formatos reais coexistem:

```
Formato A (Cyberpunk):  <li><strong>Processor:</strong> Core i7-12700
Formato B (Terraria):   <li><strong>Processor: Dual Core 3.0 Ghz</strong>
```

Um parser que assuma "valor = texto após `</strong>`" devolve **string vazia
para todos os campos do Formato B**. A estratégia adotada extrai o **texto** de
cada `<li>` e divide no primeiro `:`, tratando A e B de forma uniforme sem
depender do markup — que é justamente a parte instável.

Rótulos são normalizados por mapa de sinônimos (`Storage`/`Hard Disk Space`,
`Graphics`/`Video Card`, ...). O `*` em `OS *` (SO legado) é capturado como
`os_legacy_flag`, não descartado. Rótulo desconhecido vai para
`unparsed_labels` — **nunca é descartado em silêncio**.

**Escopo deliberadamente limitado (item 5):** extrai texto. Não converte
`"16 GB RAM"` para `16`, não mapeia GPU para benchmark. A normalização de
hardware pertence ao Data Wrangling; aqui a prioridade é fidelidade.

O `raw` HTML íntegro é **sempre** preservado, permitindo corrigir o parser
depois sem re-raspar a Steam.

---

## Robustez

- **Checkpoint por app_id** com escrita atômica (`.tmp` + `os.replace`, com
  retry para bloqueios transitórios do Windows/OneDrive).
- **Idempotente:** app concluído é pulado **sem requisição de rede**.
- **`Ctrl+C`** faz flush e encerra com resumo; retomar continua no app seguinte.
- **Falha individual nunca interrompe o lote** — cada app roda isolado.
- **Erros classificados:** transitório (timeout, 429, 5xx, JSON inválido) faz
  retry com backoff; permanente (404, 403, age gate por login) não.
- **Distinção do item 8** em campos separados: `collection_status` (dado ausente
  vs. página inacessível vs. falha HTTP vs. falha de parsing) e
  `exclusion_reason` (exclusão metodológica). Nunca confundidos.

Um jogo sem requisitos recomendados **não é erro**:
`has_recommended_requirements=false`, campos `null`, validação aprova.

---

## Age gate

A Steam usa duas variantes de bloqueio por idade, e o código as trata de forma
diferente:

| Variante | Comportamento | Tratamento |
|---|---|---|
| **interstitial na página** | HTTP 200, ~51 KB, sem `InitAppTagModal`, sem redirect | controlado por `accept_age_gate_interstitial`; auto-declaração de data de nascimento, **sem autenticação** |
| **redirect para login** | `/login/?redir=agecheck/...` | **nunca contornado**; `tags_source=AGE_GATE_LOGIN_REQUIRED`, requisitos ainda coletados via `appdetails` |

Nenhum CAPTCHA, autenticação ou mecanismo anti-bot é contornado. O `robots.txt`
da Steam permite `/app/` (bloqueia apenas `/share/`, `/news/externalpost/`,
URLs com token, `/login/?guestpasskey`, `/join/?redir`, `/account/ackgift/`,
`/email/` e `/widget/`).

---

## Limitações a declarar na metodologia do TCC

1. Discovery via storefront enumera só o que está **visível hoje**. Jogos
   deslistados desaparecem, e a probabilidade de deslistagem cresce com a idade
   → **viés de sobrevivência correlacionado com o tempo**, que é o eixo do
   estudo. Tende a enviesar *para cima* a taxa de crescimento estimada.
   Use `audit-bias` para quantificar.
2. Tags Steam são atribuídas **por usuários**, não pelo desenvolvedor. São
   ruidosas por construção: 1.915 jogos têm simultaneamente `3D` e `2D`.
3. Tags refletem o estado **atual**, não o do lançamento — anacronismo
   estrutural em estudo longitudinal.
4. Requisitos refletem a versão **atual** da página, não a do lançamento.
   Patches e remasters alteram os requisitos publicados, sem histórico acessível.
5. Requisitos são declarados pelo desenvolvedor, sem padronização nem auditoria.
6. `Indie` exclui 48,4% da população 3D e **não mede orçamento** (decisão D2).
7. `review_count` usa `appdetails.recommendations.total`, que **difere** de
   `appreviews.total_reviews` (CS2: 5.245.985 vs 9.826.091). Nunca comparar as
   duas escalas.
8. Jogos com age gate por login têm tags inacessíveis.
9. Rate limits são empíricos desta data e podem mudar.

---

## Documentação

| Documento | Conteúdo |
|---|---|
| **`docs/FASE5_guia_execucao.md`** | **guia passo a passo para executar a coleta completa** |
| `docs/FASE1_investigacao.md` | investigação dos endpoints, casos de borda, funil medido |
| `docs/FASE2_arquitetura.md` | arquitetura, fluxo, schema, checkpoints, erros |
| `docs/DECISOES.md` | decisões metodológicas D1–D13, com a evidência de cada uma |
