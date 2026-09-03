# FASE 1 — Relatório de Investigação das Fontes de Dados da Steam

**Data da investigação:** 2026-08-30
**Executado por:** scripts em `investigation/` (reprodutíveis)
**Amostras brutas:** `investigation/samples/`
**Método:** sondagem empírica de 60+ requisições reais. Nenhum endpoint foi
assumido como funcional a partir de documentação histórica ou de conhecimento
prévio; todos foram testados nesta data.

Scripts:

| Script | Objetivo |
|---|---|
| `probe_endpoints.py` | Sondagem inicial de 19 endpoints/casos |
| `analyze_samples.py` | Dissecação dos payloads + variantes de discovery |
| `probe_round3.py` | Tag IDs reais, age gate, paginação profunda |
| `probe_round4.py` | Semântica AND/OR do search, funil por tag |
| `probe_round5.py` | Conclusão do funil, `untags`, limites de taxa |

---

## 1. Achados que contradizem premissas usuais

Três resultados alteram materialmente o desenho do pipeline:

### 1.1. `ISteamApps/GetAppList/v2` está MORTO (HTTP 404)

Este é o endpoint canônico de discovery citado na maioria dos tutoriais. Foram
testadas cinco variantes, **todas falharam**:

| Variante | Resultado |
|---|---|
| `api.steampowered.com/ISteamApps/GetAppList/v2/` | **HTTP 404** |
| `.../GetAppList/v2` (sem barra) | HTTP 404 |
| `.../GetAppList/v0002/?format=json` | HTTP 404 |
| `.../GetAppList/v1/` | HTTP 404 |
| `.../ISteamApps/GetAppList/` | HTTP 404 |
| `api.steampowered.com/IStoreService/GetAppList/v1/` | **HTTP 403** (exige API key) |
| `store.steampowered.com/actions/GetAppList/?appType=game` | HTTP 200, mas devolve **página HTML "Welcome to Steam"**, não JSON |

**Consequência metodológica:** não existe, sem API key, um caminho gratuito para
enumerar o catálogo completo da Steam. As duas rotas viáveis são:

- **(A)** obter uma Steam Web API key e usar `IStoreService/GetAppList/v1`
  (paginado, suporta `include_games`, `last_appid`);
- **(B)** enumerar via `store.steampowered.com/search/results/`, que **já
  suporta filtragem por tag e por SO no servidor**.

A rota (B) é atrativa porque empurra a filtragem para o lado da Steam, reduzindo
drasticamente o volume de `appdetails` a buscar. Mas introduz um viés que precisa
ser declarado — ver §6.1.

### 1.2. O `appdetails` NÃO retorna as Steam tags

Verificado nos 9 jogos amostrados: nenhum possui campo `tags` ou `steam_tags`.
As chaves de topo são estáveis e incluem `genres` e `categories`, que **não são
tags**. Exemplo real (Cyberpunk 2077):

- `genres` = `['RPG']` — apenas 1 gênero
- `categories` = `['Single-player', 'Steam Achievements', ...]` — recursos, não conteúdo
- as 20 tags reais (`Cyberpunk`, `Open World`, `FPS`, `First-Person`,
  `Immersive Sim`, `Action RPG`, ...) **só existem no HTML da página**

**Consequência:** toda a metodologia de filtragem por tags dos itens 2.5–2.13 do
projeto **depende de uma fonte que o `appdetails` não fornece**. Existem duas
fontes de tags:

| Fonte | Tags entregues | Custo |
|---|---|---|
| HTML da página do jogo, via `InitAppTagModal(...)` | **20 tags com `tagid` + `name`** | 1 requisição (~130–270 KB) |
| `data-ds-tagids` nas linhas do `/search/results/` | **apenas 7–8 tagids** (subconjunto) | grátis, junto do discovery |

Medição direta: para Cyberpunk 2077 o HTML entrega 20 tags; o search entrega 7.
Confirmei que o conjunto do search é **subconjunto** do conjunto do HTML. Usar o
search como única fonte de tags produziria **falsos negativos** — o que o item
2.16 do projeto (priorizar recall) proíbe explicitamente.

### 1.3. Os tag IDs comumente citados estão errados

Minha primeira sondagem usou `tags=4166` supondo ser "3D". **4166 é
`Atmospheric`.** O ID real de `3D` é **4191**. Isso invalidaria silenciosamente
todo o filtro. Endpoint de resolução confirmado funcional:

- `store.steampowered.com/actions/ajaxgetstoretags` → 430 tags
- `api.steampowered.com/IStoreService/GetTagList/v1/?language=english` → **446 tags, sem API key**
- `store.steampowered.com/tagdata/populartags/english` → 430 tags

O pipeline **deve resolver nome→tagid dinamicamente** e falhar de forma ruidosa
se uma tag do `filters.yaml` não existir na taxonomia oficial.

---

## 2. Endpoints validados (a usar)

### 2.1. Taxonomia de tags

```
GET https://api.steampowered.com/IStoreService/GetTagList/v1/?language=english
```

- **Auth:** nenhuma. **HTTP 200**, 15.792 bytes, 446 tags.
- **Resposta:** `{"response": {"tags": [{"tagid": 9, "name": "Strategy"}, ...]}}`
- **Uso:** resolver os nomes de tags do `filters.yaml` para IDs; versionar o
  snapshot da taxonomia para reprodutibilidade.

Resolução das tags definidas no projeto (§2.6–2.13):

| Tag do projeto | tagid | Papel no filtro |
|---|---|---|
| **3D** | **4191** | positiva forte (núcleo) |
| First-Person | 3839 | positiva forte |
| Third Person | 1697 | positiva forte |
| Third-Person Shooter | 3814 | positiva forte |
| FPS | 1663 | positiva forte |
| 3D Platformer | 5395 | positiva forte |
| 3D Fighter | 6506 | positiva forte |
| Immersive Sim | 9204 | positiva forte |
| Walking Simulator | 5900 | positiva forte |
| Automobile Sim | 1100687 | positiva forte |
| Flight | 15045 | positiva forte |
| Space Sim | 16598 | positiva forte |
| Looter Shooter | 353880 | positiva forte |
| Hero Shooter | 620519 | positiva forte |
| Arena Shooter | 5547 | positiva forte |
| **Rail Shooter** | **3954** | positiva forte — **0 jogos, ver §5** |
| Open World | 1695 | positiva secundária |
| Realistic | 4175 | positiva secundária |
| Cinematic | 4145 | positiva secundária |
| Action-Adventure | 4106 | positiva secundária |
| Action RPG | 4231 | positiva secundária |
| Souls-like | 29482 | positiva secundária |
| Survival Horror | 3978 | positiva secundária |
| Driving | 1644 | positiva secundária |
| 2D | 3871 | negativa |
| 2D Platformer | 5379 | negativa |
| 2D Fighter | 4736 | negativa |
| 2.5D | 4975 | **neutra** — nunca excluir (§2.9) |
| Indie | 492 | negativa (revisável) |
| Pixel Graphics | 3964 | negativa candidata |
| Hand-drawn | 6815 | negativa candidata |
| Side Scroller | 3798 | negativa candidata |
| Text-Based | 31275 | negativa candidata |
| Visual Novel | 3799 | negativa candidata |
| Interactive Fiction | 11014 | negativa candidata |
| Point & Click | 1698 | negativa candidata |
| Card Game | 1666 | negativa candidata |
| Board Game | 1770 | negativa candidata |
| Hidden Object | 1738 | negativa candidata |
| Casual | 597 | **nunca usar como exclusão** (§2.12) |
| Free to Play | 113 | **nunca usar como exclusão** (§2.12) |
| Early Access | 493 | **nunca usar como exclusão simples** (§2.12) |
| ~~Homemade~~ | **inexistente** | **confirmada a suspeita do item 2.13** |

**"Homemade" não existe na taxonomia oficial de 446 tags.** O item 2.13 do
projeto pedia confirmação antes de criar a regra — está confirmado: a regra
`NOT Homemade` não deve ser criada.

### 2.2. Discovery filtrado

```
GET https://store.steampowered.com/search/results/
    ?query=&start=0&count=50&infinite=1&json=1
    &category1=998        # 998 = Games (exclui DLC/software/vídeo no servidor)
    &os=win               # exige suporte a Windows
    &tags=4191            # tag(s)
    &untags=492           # exclusão de tag(s)
    &sort_by=Released_DESC
```

Resposta: `{"success", "total_count", "start", "results_html"}`.

Achados críticos de comportamento:

- **`tags` com múltiplos valores é AND, não OR.** Medido: `3D`=51.760,
  `FPS`=10.988, `3D,FPS`=6.503. Como 6.503 ≤ min(51.760, 10.988), a semântica é
  interseção. **Isso significa que a união das tags positivas do projeto (§2.6)
  não pode ser feita em uma única query** — exige uma query por tag e união do
  lado do cliente.
- **`untags` é complemento exato.** Validado por soma: `3D AND Indie`=25.036 +
  `3D NOT Indie`=26.724 = 51.760 = `3D`. Diferença zero. Exclusão confiável no
  servidor.
- **Paginação profunda funciona.** `start` testado em 0 / 1.000 / 5.000 / 20.000
  / 38.000 / 45.000 — todos HTTP 200 com 50 resultados. Sem teto observado.
- **`results_html` traz `data-ds-appid` e `data-ds-tagids`** por linha, mas
  apenas 7–8 tagids (subconjunto — ver §1.2).
- **Não expõe filtro por intervalo de data de lançamento.** O corte de 2005
  (§2.3) precisa ser aplicado após o `appdetails`.

### 2.3. Metadados + requisitos de hardware (endpoint principal)

```
GET https://store.steampowered.com/api/appdetails?appids=<id>&l=english&cc=us
```

- **Auth:** nenhuma. JSON limpo. `{"<appid>": {"success": bool, "data": {...}}}`
- **`l=english` é obrigatório** para estabilidade do parsing: os rótulos dos
  requisitos (`OS:`, `Processor:`, `Memory:`, `Graphics:`, `Storage:`) vêm
  localizados. Sem fixar o idioma, o parser quebraria de forma inconsistente.
- **`cc=us` fixa a moeda/região**, evitando variação regional nos metadados.

Campos confirmados presentes e úteis:

| Campo | Uso | Confirmado em |
|---|---|---|
| `type` | filtro §2.1 (`game` vs `dlc`) | discrimina corretamente: app 2138330 = `"dlc"` |
| `name`, `steam_appid` | identificação | todos |
| `release_date.date` | filtro §2.3 (formato `"Dec 9, 2020"`) | todos |
| `release_date.coming_soon` | filtro §2.4 (estado de lançamento) | todos |
| `platforms.{windows,mac,linux}` | filtro §2.2 | todos |
| `genres`, `categories` | classificação (não são tags) | todos |
| `developers`, `publishers` | proxy futura de escala (§2.15) | todos |
| `supported_languages` | proxy futura de escala (§2.15) | todos (string HTML) |
| **`recommendations.total`** | **review_count grátis** | todos |
| `metacritic.score` | metadado auxiliar | 7 de 9 |
| `required_age`, `is_free` | auxiliares | todos |
| **`pc_requirements`** | **objetivo central** | ver §3 |
| `linux_requirements`, `mac_requirements` | ignorar (fora do escopo) | todos |

**`recommendations.total` elimina a necessidade de uma chamada separada de
reviews** para o filtro do §2.14. Um único `appdetails` entrega metadados +
requisitos + contagem de reviews. Isso reduz o custo do scraping pela metade.

### 2.4. Reviews (opcional, apenas se o split positivo/negativo for necessário)

```
GET https://store.steampowered.com/appreviews/<id>?json=1&num_per_page=0
```

Retorna somente o sumário (~200 bytes), sem baixar reviews individuais:
`{"num_reviews":0, "review_score":8, "review_score_desc":"Very Positive",
"total_positive":8442153, "total_negative":1383938, "total_reviews":9826091}`

Nota de divergência: para CS2, `appreviews.total_reviews`=9.826.091 mas
`appdetails.recommendations.total`=5.245.985. **São métricas diferentes** —
`recommendations` conta apenas reviews em determinado escopo de idioma/tipo de
compra. Se o threshold do §2.14 for aplicado, é preciso decidir qual métrica é a
oficial e usá-la consistentemente. **Questão metodológica aberta — ver §7.**

### 2.5. HTML da página do jogo (necessário exclusivamente para as 20 tags)

```
GET https://store.steampowered.com/app/<id>/
```

Tags extraíveis de `InitAppTagModal(<appid>, [{"tagid":...,"name":...}, ...],`
— 20 tags com ID e nome. Também contém `game_area_sys_req`, mas os requisitos
já vêm melhor estruturados no `appdetails`; **não farei parsing de requisitos do
HTML** (respeitando a ordem de prioridade do item 3 do projeto).

---

## 3. `pc_requirements` — estrutura real e riscos de parsing

Estrutura: `{"minimum": "<html>", "recommended": "<html>"}`. Ambos são **strings
HTML**, não objetos. `recommended` **pode estar ausente**.

Contagem na amostra: **6 de 9 jogos têm `recommended`; 3 não têm**
(CS2, Half-Life 2, Portal 2 — todos da Valve). Isso valida a exigência do item 8
do projeto: ausência de recomendados **não pode** derrubar o pipeline.

### 3.1. Ao menos DOIS formatos de markup distintos coexistem

**Formato A — canônico** (`<li><strong>Rótulo:</strong> valor`), Cyberpunk 2077:

```html
<strong>Recommended:</strong><br><ul class="bb_ul">
<li>Requires a 64-bit processor and operating system<br></li>
<li><strong>OS:</strong> 64-bit Windows 10<br></li>
<li><strong>Processor:</strong> Core i7-12700 or Ryzen 7 7800X3D<br></li>
<li><strong>Memory:</strong> 16 GB RAM<br></li>
<li><strong>Graphics:</strong> GeForce RTX 2060 SUPER or Radeon RX 5700 XT or Arc A770<br></li>
<li><strong>DirectX:</strong> ...
```

**Formato B — rótulo E valor dentro do `<strong>`**, Terraria:

```html
<h2 class="bb_tag"><strong>RECOMMENDED</strong></h2><ul class="bb_ul">
<li><strong>OS: Windows 7, 8/8.1, 10</strong> <br></li>
<li><strong>Processor: Dual Core 3.0 Ghz</strong> <br></li>
<li><strong>Memory: 4GB</strong><br></li>
<li><strong>Hard Disk Space: 200MB </strong><br></li>
<li><strong>Video Card: 256mb Video Memory, capable of Shader Model 2.0+</strong>
```

Um parser que assuma "valor = texto após o `</strong>`" retorna **string vazia
para todos os campos do Formato B**. Um parser que assuma o cabeçalho
`<strong>Recommended:</strong>` não reconhece `<h2 class="bb_tag">`.

### 3.2. Rótulos não são estáveis entre jogos

Observados para o mesmo conceito:

- Armazenamento: `Storage:`, `Hard Disk Space:`
- GPU: `Graphics:`, `Video Card:`
- SO: `OS:` e **`OS *:`** (o asterisco marca SO legado sem suporte — Half-Life 2,
  Portal 2, Hades)

O parser precisa de um **mapa de sinônimos de rótulos**, e o `*` deve ser
capturado como sinal (`os_unsupported_flag`), não descartado.

**Decisão de projeto:** o `raw` HTML íntegro de `minimum` e `recommended` será
sempre persistido, exatamente como exigido pelo item 5 do projeto. Qualquer erro
de parsing será corrigível offline, sem re-raspar a Steam.

---

## 4. Casos de borda — comportamento medido

| Caso | Teste | Comportamento real |
|---|---|---|
| **app_id inexistente** | 999999999 | HTTP **200** com `{"999999999":{"success":false}}`. **Não é 404.** O código de status é inútil como sinal de erro; é obrigatório checar `success`. |
| **DLC** | 2138330 (Phantom Liberty) | `type="dlc"`, mas **possui `pc_requirements` completos, idênticos ao jogo-base**. Sem o filtro `type=="game"` o dataset seria contaminado por duplicatas de requisitos. Confirma o item 2.1. |
| **Age gate — maioria** | GTA V (271590), Manhunt (12130), Postal 2 (232770), RDR2 (1174180) | Sem cookie de idade: HTTP 200, página **completa**, com `InitAppTagModal` e `game_area_sys_req`. Age gate **não bloqueia** jogos mature. |
| **Age gate — real** | HuniePop (339800) | Redireciona para `/login/?redir=agecheck/app/339800/`. **Exige login, não apenas cookie.** Página de tags inacessível. |
| **Jogo pré-2005** | Half-Life 2 (2004) | `release_date.date="Nov 16, 2004"` → excluído corretamente pelo §2.3, com a data original preservada. |
| **Sem requisitos recomendados** | CS2, HL2, Portal 2 | `pc_requirements` sem a chave `recommended`. Sem exceção; tratar como `null` + `has_recommended_requirements=false`. |
| **F2P** | CS2 | Processado normalmente. Confirma §2.12 (não excluir F2P). |
| **Rótulo `OS *`** | HL2, Portal 2, Hades | SO legado marcado com asterisco. |

**Sobre o age gate com login:** não tentarei contornar. Jogos nessa situação
serão registrados com `exclusion_reason=AGE_GATE_LOGIN_REQUIRED` e, se o
`appdetails` responder (a API costuma responder mesmo quando o HTML não), os
requisitos serão coletados e apenas as tags ficarão ausentes
(`tags_source=UNAVAILABLE`). Isso preserva o dado principal sem violar bloqueios.

---

## 5. Funil real medido (`samples/_funnel_counts.json`)

Baseline: **172.017 jogos** com suporte a Windows na loja (`category1=998&os=win`).

Tags positivas fortes (§2.6):

| Tag | Jogos | | Tag | Jogos |
|---|---:|---|---|---:|
| 3D | 51.760 | | Arena Shooter | 4.196 |
| First-Person | 30.590 | | Flight | 3.002 |
| Third Person | 18.510 | | 3D Fighter | 2.165 |
| FPS | 10.988 | | Automobile Sim | 2.211 |
| Immersive Sim | 9.954 | | Space Sim | 2.083 |
| 3D Platformer | 9.875 | | Looter Shooter | 1.547 |
| Walking Simulator | 7.993 | | Hero Shooter | 1.339 |
| Third-Person Shooter | 4.617 | | **Rail Shooter** | **0** |

**`Rail Shooter` (3954) retorna 0 jogos.** A tag existe na taxonomia oficial mas
não tem nenhum jogo com suporte a Windows associado no search. Provavelmente é
uma tag obsoleta/não-curada. Vou mantê-la no `filters.yaml` (é inócua) mas
registrar em log que ela não contribui com nenhum candidato.

Negativas e neutras:

| Tag | Jogos | Observação |
|---|---:|---|
| Indie | **95.939** | 56% de toda a loja |
| Casual | 74.127 | nunca usar como exclusão |
| **2D** | **63.469** | maior que 3D (51.760) |
| Pixel Graphics | 32.903 | |
| 2D Platformer | 14.965 | |
| Visual Novel | 11.982 | |
| Side Scroller | 8.657 | |
| **2.5D** | **5.922** | nunca excluir |
| 2D Fighter | 2.759 | |

Interseções decisivas para a metodologia:

| Recorte | Jogos | Leitura |
|---|---:|---|
| 3D | 51.760 | universo da tag núcleo |
| **3D ∩ Indie** | **25.036** | **48,4% de todos os jogos 3D são Indie** |
| 3D \ Indie | 26.724 | efeito do §2.10 |
| 3D \ (Indie ∪ 2D) | 25.575 | |
| **3D ∩ 2D** | **1.915** | **conflito real do §2.8 confirmado** |
| 3D ∩ 2.5D | 1.071 | híbridos do §2.9 |
| 3D ∩ Pixel Graphics | 2.389 | confirma o alerta do §2.11 |

Três confirmações empíricas das cautelas do projeto:

1. **§2.8 estava certo:** 1.915 jogos têm simultaneamente `3D` e `2D`. Uma
   exclusão rígida `NOT 2D` descartaria silenciosamente 1.915 jogos 3D. Serão
   marcados `tag_conflict_3d_2d=true` e preservados para resolução posterior.
2. **§2.11 estava certo:** 2.389 jogos são `3D` **e** `Pixel Graphics`. Nenhuma
   das negativas candidatas será aplicada como exclusão rígida — apenas marcada.
3. **§2.10 é a decisão de maior impacto do estudo:** excluir `Indie` corta
   **48,4% da população 3D**. Precisa de decisão explícita — ver §7.

---

## 6. Limites de taxa e viés — medidos

### 6.1. Viés de sobrevivência no discovery via search (metodologicamente relevante)

O `/search/results/` retorna apenas o que está **atualmente visível na
storefront**. Jogos deslistados (retirados de venda, licenças de música
expiradas, remoções por editora) **não aparecem**. Como a probabilidade de
deslistagem cresce com a idade do título, isso produz **viés de sobrevivência
correlacionado com o tempo** — exatamente o eixo do estudo longitudinal.

Efeito esperado: sub-representação de jogos de 2005–2012, o que pode **enviesar
para cima** a estimativa da taxa de crescimento dos requisitos (os jogos antigos
que sobrevivem na loja tendem a ser os mais bem-sucedidos, e portanto os de maior
orçamento e requisitos mais altos).

A rota com API key (`IStoreService/GetAppList`) enumera o catálogo por app_id e é
menos suscetível a esse viés. **Requer decisão — ver §7.**

### 6.2. Rate limiting medido empiricamente

| Endpoint | Ritmo testado | Resultado |
|---|---|---|
| `/search/results/` | ~1,0 s/req | **HTTP 429 na 26ª requisição** |
| `/search/results/` | ~0,35 s/req | **HTTP 429 na 17ª requisição, após 12,4 s** |
| `/api/appdetails` | ~0,35 s/req | **30/30 HTTP 200 em 19,0 s, sem 429** |

- **Nenhum header informativo:** `Retry-After` **ausente**, nenhum header
  `X-RateLimit-*`. O backoff precisa ser cego (exponencial com jitter).
- `/search/results/` é **muito** mais restrito que `/api/appdetails`. Precisa de
  budget próprio, mais lento (alvo: ~1 req / 3–5 s).
- Recuperação após 429 confirmada: backoff exponencial de 3 s → 6 s → 12 s
  restabeleceu as respostas 200 nas rodadas 4 e 5.

**Estimativa de custo do scraping.** Um candidato = 1 `appdetails` (+1 HTML se
tags completas forem necessárias). Para ~26.700 candidatos (3D \ Indie):

| Cenário | Requisições | Tempo estimado @0,5 s/req |
|---|---:|---:|
| Só `appdetails` (tags do search, 7 tagids) | ~26.700 | **~3,7 h** |
| `appdetails` + HTML (20 tags, recall alto) | ~53.400 | **~7,4 h** |

Ambos são viáveis com checkpoint/resume. A diferença é qualidade de recall das
tags vs. tempo de máquina.

---

## 7. Questões metodológicas que exigem SUA decisão

Conforme o item 15 do projeto, não vou resolvê-las arbitrariamente. São quatro, e
todas alteram a população do estudo:

**Q1 — Estratégia de discovery e viés de sobrevivência.**
Search filtrado (rápido, sem credenciais, mas com viés de sobrevivência
correlacionado ao tempo — §6.1) vs. Steam Web API key
(`IStoreService/GetAppList`, enumera o catálogo, mitiga o viés, mas exige uma
key gratuita e ~35× mais requisições de `appdetails`)?

**Q2 — Exclusão de `Indie` (impacto: 48,4% da população 3D).**
Excluir na fonte (menor custo, mas descarta 25.036 jogos 3D — e a tag `Indie` na
Steam é auto-declarada, marcando títulos como Hades, Hollow Knight e Valheim) vs.
apenas marcar `is_indie=true` e coletar tudo, decidindo a exclusão no
Data Wrangling com os dados na mão?

**Q3 — Fonte das tags (impacto: recall vs. 2× o tempo).**
7–8 tagids grátis do search (rápido, risco de falso negativo, contra o §2.16) vs.
20 tags completas do HTML de cada jogo (dobra o tempo, recall alto)?

**Q4 — Métrica oficial de contagem de reviews.**
`appdetails.recommendations.total` (grátis, mas para CS2 dá 5.245.985) vs.
`appreviews.total_reviews` (+1 requisição por jogo, para CS2 dá 9.826.091). São
métricas diferentes; o threshold do §2.14 depende de qual for adotada.

---

## 8. Limitações conhecidas a declarar na metodologia do TCC

1. `ISteamApps/GetAppList/v2` descontinuado; sem API key não há enumeração
   gratuita do catálogo completo (§1.1).
2. Discovery via storefront introduz viés de sobrevivência correlacionado ao
   tempo (§6.1).
3. Tags Steam são **atribuídas por usuários**, não pelo desenvolvedor. São
   ruidosas por construção: 1.915 jogos são simultaneamente `3D` e `2D`.
4. Tags refletem o **estado atual**, não o do lançamento. Um jogo de 2006 tem
   tags aplicadas por usuários em 2015+ — anacronismo estrutural em estudo
   longitudinal.
5. Requisitos publicados refletem a **versão atual** da página, não a do
   lançamento. Jogos com patches/remasters podem exibir requisitos revisados,
   sem histórico acessível.
6. Requisitos são declarados pelo desenvolvedor, sem padronização nem auditoria.
   Formatos variam (§3.1) e valores podem ser aspiracionais.
7. Jogos com age gate por login (ex.: HuniePop) têm tags inacessíveis sem
   autenticação; não serão contornados.
8. `Rail Shooter` (3954) existe na taxonomia mas retorna 0 jogos (§5).
9. `Homemade` **não existe** na taxonomia oficial (§2.1).
10. `recommendations.total` ≠ `appreviews.total_reviews` (§2.4).
11. Rate limiting não documentado, sem `Retry-After`; os limites de §6.2 são
    empíricos desta data e podem mudar.
