# Controle de Demandas — Documentação

> App de controle de demandas (orçamentos e revisões) de uma empresa de esquadrias de alumínio.
> Arquivo único `Controle_Demandas_Kanban.html`. Offline, autossuficiente, sem dependências de servidor.

---

## Índice

1. [Visão geral](#1-visão-geral)
2. [Manual do usuário](#2-manual-do-usuário)
3. [Documentação técnica](#3-documentação-técnica)
4. [Decisões e armadilhas conhecidas](#4-decisões-e-armadilhas-conhecidas)
5. [Como modificar / estender](#5-como-modificar--estender)
6. [Solução de problemas](#6-solução-de-problemas)
7. [Publicação no GitHub](#7-publicação-no-github)

---

## 1. Visão geral

### O que é
Um *planner* pessoal em formato Kanban para acompanhar as demandas que chegam dos vendedores (orçamentos, revisões, fechamentos MoB, atualizações). Cada cartão representa **uma demanda** e passa por três fases: **Não iniciado → Na Revisão → Concluído**.

### Quem usa
Um único usuário, num único PC. Não foi projetado para uso compartilhado.

### O que substituiu
Uma planilha `.xlsm` com formulário de cadastro feito em VBA. A planilha original tinha ~1300 linhas (set/2024 a mai/2026), todas migradas para o app na primeira versão.

### Stack
- **Arquivo único HTML** (~350 KB) com CSS e JavaScript embutidos.
- **Sem frameworks, sem build, sem servidor.**
- Fontes via Google Fonts com *fallback* nativo do sistema (funciona offline mesmo sem internet).
- Logo da marca embutida como base64.

### Como executar
Clique duplo no `.html` — abre no navegador padrão. Recomendado: **Chrome ou Edge** (o vínculo de arquivo de backup só funciona nesses dois).

---

## 2. Manual do usuário

### 2.1 As três abas

**Quadro** (padrão ao abrir) — o Kanban com as três colunas.

**Painel** — indicadores e tabela por vendedor.

**Vendedores** — cadastro de quem manda demanda pra você.

### 2.2 Quadro Kanban

#### Colunas
| Coluna | Significado |
|---|---|
| **Não iniciado** | demandas que entraram e ainda não comecei |
| **Na Revisão** | em revisão / em andamento |
| **Concluído** | entregue (histórico) |

#### Como arrastar
- Segura no cartão e arrasta para a coluna desejada.
- Ao soltar em **Na Revisão**, a data de revisão é preenchida automaticamente com a data de hoje (se estava vazia).
- Ao soltar em **Concluído**, a data de entrega é preenchida automaticamente (se estava vazia).
- Se a data já estiver preenchida, o app **não sobrescreve** — você precisa editar manualmente.

#### O que cada cartão mostra
- Badge do tipo (Orçamento / Revisão / MoB / Atualização)
- Número da demanda (`#1234`)
- Cliente / obra
- Avatar e nome do vendedor
- **Idade em dias** (só para cartões ativos) — amarelo a partir de 3 dias, vermelho a partir de 7
- As três datas no rodapé do cartão: Ent / Rev / Entr

#### Coluna Concluído (histórico)
- Mostra apenas os 20 mais recentes (mais novos no topo).
- Botão **"Carregar mais (+20)"** no fim da lista.
- Ao buscar ou filtrar, **todos** os concluídos aparecem normalmente.

#### Busca e filtros
- **Campo de busca** — varre cliente, vendedor, tipo e número em todas as colunas simultaneamente.
- **Filtro por vendedor** — ativos aparecem no topo da lista.
- **Filtro por tipo** — Orçamento / Revisão / MoB / Atualização.

#### Criar / editar / excluir demanda
- Botão **"+ Nova demanda"** — abre modal com cliente, vendedor (com sugestões), tipo, status e datas.
- **Clicar em um cartão** abre o mesmo modal em modo edição, com botão de excluir.
- Se você digitar um nome de vendedor que ainda **não existe** no cadastro, ele é criado automaticamente como ativo.
- ID é sequencial automático; data de entrada padrão é hoje.

### 2.3 Painel

#### Indicadores no topo
- **Demandas ativas** — não concluídas (total)
- **Concluídas neste mês** — entregues no mês corrente
- **Médio até revisão** — dias entre entrada e ida pra revisão (média do período)
- **Médio até conclusão** — dias entre entrada e entrega (média do período)

#### Tabela por vendedor
- Lista todos os vendedores com pelo menos uma demanda no período, ou que estão marcados como ativos.
- Mostra: total no período · neste mês · neste ano · médio até revisão · médio até conclusão · barra de volume.
- **Ativos** ficam no topo com a tag verde. **Inativos** ficam abaixo com a tag cinza.

#### Gráfico de volume
- Por **mês** quando um ano específico está selecionado.
- Por **ano** quando "Todos os anos" está selecionado.

#### Filtros do painel
- **Período**: ano específico ou "Todos os anos".
- **Tipo**: por padrão "todos os tipos", pra enxergar orçamento, revisão e fechamento juntos.

### 2.4 Vendedores

#### O que faz
Cadastro central de quem manda demanda pra você. **Não** é a lista de cartões — é o *registro de pessoas*.

#### Operações
- **+ Novo vendedor** — abre modal com nome, telefone, e-mail e marcação de ativo.
- **Ícone de lápis** — edita o cadastro.
- **Interruptor** — ativa/desativa diretamente da tabela.
- **Ícone de lixeira** — exclui. **Só é permitido se o vendedor não tiver nenhuma demanda vinculada**. Se tiver, o app bloqueia e sugere desativar.

#### Comportamentos importantes
- **Renomear vendedor com demandas vinculadas:** o app avisa quantos cartões serão atualizados e propaga o novo nome para todos eles. Isso é intencional — evita ter "Bruno" e "Bruno Almeida" no histórico.
- **Ativo/inativo:** apenas afeta onde o vendedor aparece nas listas. Ativos vão pro topo dos filtros do quadro, da datalist do modal e da tabela do painel. **Os cartões dele continuam visíveis e contabilizados normalmente** — desativar não esconde histórico.
- **Auto-cadastro:** quando você digita um nome novo numa Nova demanda, ele é criado como ativo. Depois é só ajustar no cadastro se quiser.

### 2.5 Backup

#### Os dois estados do selo no topo

| Selo | Significado | O que acontece se eu limpar o cache do navegador |
|---|---|---|
| 🟡 **"Backup: só neste navegador"** | dados só no localStorage | **perde tudo** e volta ao seed embutido (1300 originais) |
| 🟢 **"Backup: arquivo vinculado"** | gravando em arquivo real no disco | **dados preservados** no arquivo |

#### Como vincular (uma vez só)
1. Clica no selo amarelo no topo (ou menu ⋮ → "Vincular arquivo de backup").
2. Escolhe uma pasta e um nome (`controle_demandas_backup.json` é o sugerido).
3. Pronto. O selo fica verde.

A partir daí, o app:
- Grava nesse arquivo **a cada alteração** (com pequeno *debounce* de 1.5s).
- Grava também nos **checkpoints às 12h e 17h** quando o app está aberto.
- O navegador pode pedir permissão de leitura/escrita uma vez por sessão — é normal.

#### Sem vínculo
Os checkpoints das 12h/17h tentam baixar o JSON automaticamente. Alguns navegadores bloqueiam download automático sem clique do usuário — por isso o **vínculo é o caminho recomendado**.

### 2.6 Exportação e importação

No menu ⋮:
- **Exportar p/ Excel (CSV)** — abre no Excel preservando os campos.
- **Exportar backup (JSON)** — formato completo (cartões + vendedores).
- **Importar backup (JSON)** — substitui tudo. Aceita o formato novo (objeto com `cards` e `vendors`) ou o antigo (só uma lista de cartões).

### 2.7 Tema claro/escuro

Botão de sol/lua no cabeçalho:
- Alterna entre claro e escuro.
- A escolha fica **salva** — abre sempre no que você deixou.
- Na primeira vez, **respeita o tema do sistema** (se o Windows está no modo escuro, abre escuro).

### 2.8 Restaurar dados originais

Menu ⋮ → "Restaurar dados originais". Volta para os **1300 cartões embutidos** da planilha original. **Não pode ser desfeito** — faça um backup antes se quiser.

---

## 3. Documentação técnica

### 3.1 Modelo de dados

#### Cartão (demanda)
```js
{
  id: 1234,                  // sequencial automático
  vendedor: "Mariana Costa",
  tipo: "Orçamento",         // | "Revisão" | "MoB - Fechamento" | "Atualização"
  cliente: "OBRA TAL",
  entrada: "2026-05-15",     // ISO (aaaa-mm-dd) ou null
  revisao: "2026-05-19",
  entrega: null,
  status: "Na Revisão"       // | "Não iniciado" | "Concluído"
}
```

#### Vendedor
```js
{
  id: 1,
  nome: "Mariana Costa",
  tel: "(15) 99999-9999",
  email: "mariana@empresa.com",
  ativo: true
}
```

#### Snapshot de backup (formato v2)
```js
{ v: 2, cards: [...], vendors: [...] }
```

### 3.2 Persistência

| Chave | Onde | Conteúdo |
|---|---|---|
| `controle_demandas_v1` | localStorage | array de cartões |
| `controle_vendedores_v1` | localStorage | array de vendedores |
| `cd_theme` | localStorage | `"light"` ou `"dark"` |
| `cd_names_v2` | localStorage | flag de migração de nomes (executada uma vez) |
| `cd_bk12`, `cd_bk17` | localStorage | data do último checkpoint 12h/17h |
| `handle` | IndexedDB `cd_backup` | FileSystemFileHandle do arquivo de backup vinculado |

### 3.3 Arquitetura do código

Todo o JavaScript fica num único `<script>` no fim do `<body>`. Organização em blocos comentados:

```
const SEED / NAME_MAP / SEED_ACTIVE        // dados embutidos + lista canônica
const KEY / VKEY / TIPOS / COLS / MONTHS   // constantes

let state / vendors                         // estado vivo (mutável)

// data layer
load() / loadVendors() / buildVendorsFromData() / migrateNames()
persist() / persistVendors()
ensureVendor() / vendCardCount() / activeNames() / isActiveVend()

// helpers
todayISO() / fmtBR() / daysBetween() / daysSince() / initials() / colorFor() / esc()
nextId() / nextVendId()

/* ============ KANBAN ============ */
matches() / renderBoard() / cardEl() / moveCard() / renderStats()

/* ============ DASHBOARD ============ */
avg() / fmtDays() / years() / renderDash()

/* ============ VIEW TOGGLE ============ */
setView() / rerender()

/* ============ VENDEDORES ============ */
renderVend() / openVend() / closeVend() / saveVend() / delVend() / toggleVend()

/* ============ MODAL (demanda) ============ */
openModal() / closeModal() / saveModal() / delCard()

/* ============ BACKUP / AUTO-SAVE ============ */
idb() / initBackup() / linkBackup() / ensurePerm() / writeBackup() / scheduleBackup() / updateBadge()
checkSchedule() / doScheduled()
snapshot() / applySnapshot()

/* ============ EXPORT / IMPORT ============ */
expCsv() / expJson() / impJson() / download() / resetData()

/* ============ THEME ============ */
applyTheme() / initTheme() / toggleTheme()

// toast / refreshFilters / event wiring no fim
```

### 3.4 Sistema de tema (CSS)

Tudo o que muda entre temas é uma **variável CSS** definida em `:root` (claro) e sobrescrita em `[data-theme="dark"]`. Para o tema funcionar em algo novo, basta usar a variável correta.

#### Tokens principais
| Token | Uso |
|---|---|
| `--paper` | fundo da página |
| `--card` | fundo de cards, modal, painel |
| `--panel` | fundo de superfícies intermediárias |
| `--ink` | texto principal |
| `--muted` | texto secundário |
| `--faint` | texto terciário/desabilitado |
| `--line`, `--line2` | divisores |
| `--accent` | cor de ação (vermelho da marca) |
| `--navy`, `--red` | cores da identidade visual da marca |
| `--todo` / `--rev` / `--done` | cores das três fases |
| `--col-bg` | fundo translúcido das colunas |
| `--av-bg`, `--av-ink` | avatar de vendedor |
| `--inp-bg`, `--menu-bg`, `--iact-bg`, `--th-bg` | superfícies de formulário, menu, ícones, cabeçalho de tabela |
| `--logo-filter` | filtro CSS aplicado à logo (inverte cores no escuro) |
| `--switch-off` | cor do switch quando desligado |

#### Anti-flash
Um *script* mínimo no `<head>` aplica `data-theme="dark"` antes do CSS pintar, lendo o localStorage ou `prefers-color-scheme`. Sem isso o app dá um flash branco ao carregar no escuro.

### 3.5 Identidade visual da marca

- **Navy** `#2e3850` — estrutura, títulos
- **Vermelho** `#c8202a` — exclusivamente ações pequenas (botões, aba ativa)
- **Cinza** `#707070` — texto auxiliar
- Tipografia: **Bricolage Grotesque** (títulos) + **IBM Plex Sans** (corpo) + **IBM Plex Mono** (números)

---

## 4. Decisões e armadilhas conhecidas

### 4.1 Por que arquivo HTML único e não app/PWA/site
- Usuário único, máquina única, sem necessidade de sincronização.
- Funciona offline.
- "Funciona sempre" — não depende de manter Node, banco, hosting.
- Backup é um arquivo texto que sobrevive a qualquer coisa.

### 4.2 Por que vendedores são um registro separado, não só texto nos cartões
Antes os ativos eram uma lista *hardcoded* no código. Toda mudança exigia editar o HTML. Hoje:
- Cadastro é a **fonte da verdade** de quem é ativo.
- Renomear propaga para todos os cartões (com aviso).
- Auto-cadastro evita órfãos.

### 4.3 Por que File System Access API em vez de só localStorage
**localStorage não protege contra limpeza de cache do navegador.** Já avisei isso ao usuário e foi o ponto que motivou o vínculo de arquivo real:
- Arquivo no disco sobrevive a limpezas de cache, reinstalações de navegador, etc.
- A API existe nativamente no Chrome/Edge — não exige instalação.
- O Firefox cai em fallback (download dos backups agendados).

### 4.4 Por que `.vbar` em vez de `.bar` no painel
**Houve uma colisão de classes**. A classe `.bar` foi inicialmente usada tanto para a barra superior do app quanto para as barrinhas de volume do painel. Como a regra das barrinhas vinha depois no CSS, o `background: var(--accent)` (vermelho) sobrescrevia o fundo da barra superior, pintando-a de vermelho. A solução foi renomear as barrinhas pra `.vbar`. **Não voltar a usar `.bar` para algo que não seja a barra superior.**

### 4.5 Mapeamento de nomes (migração automática)
Histórico tinha nomes informais. A migração foi executada uma vez (flag `cd_names_v2`):

| Antigo | Canônico |
|---|---|
| Mari | Mariana Costa |
| Carlão | Carlos Henrique |
| Rafa | Rafael Souza |
| Bruninho* | **Bruno Almeida** |
| Brunão* | **Bruno Carvalho** |

**Anderson Lima é pessoa distinta** do Rafael Souza — não unificar. Os demais ex-funcionários (Diego, Felipe, Marcos, Paulo, WAGNER*) ficaram inativos no cadastro.

### 4.6 Tempo médio é calculado só com dados completos
As médias do painel só consideram cartões em que **as duas datas existem**. Por exemplo, "médio até revisão" só conta cartões que têm `entrada` E `revisao`. Isso evita falsear o número com cartões incompletos.

### 4.7 Coluna Concluído precisa de paginação
Sem paginação, 1298 cartões na coluna Concluído travam o navegador e tornam o quadro inútil. A solução é mostrar 20 por vez (mais recentes primeiro) e desativar a paginação durante busca.

### 4.8 Avatar de cartão usa variável de tema, não navy fixo
No tema escuro o navy fixo ficaria invisível contra o fundo. Por isso o avatar usa `--av-bg` / `--av-ink`, que mudam por tema.

---

## 5. Como modificar / estender

### 5.1 Adicionar um novo tipo de demanda
1. Localiza `const TIPOS = [...]` no script.
2. Adiciona a string na lista.
3. (Opcional) Cria uma classe CSS `.tipo.t-novo` com a cor desejada e edita `tipoClass()` para mapear o novo tipo.

### 5.2 Adicionar um novo campo no cartão
1. Adiciona o campo no modal HTML (`#overlay`).
2. Lê o valor em `saveModal()` e inclui no objeto `data`.
3. Mostra no `cardEl()` se quiser que apareça no Kanban.
4. **Backup já fica retrocompatível automaticamente** — o JSON é livre.

### 5.3 Adicionar um campo no cadastro de vendedor
Mesma lógica: HTML do modal `#voverlay` → `saveVend()` → opcionalmente exibe em `renderVend()`.

### 5.4 Ajustar paleta de tema escuro
Todas as cores do escuro estão em `[data-theme="dark"] { ... }` no topo do CSS. Mexer ali atinge o app inteiro.

### 5.5 Adicionar um indicador novo no painel
Em `renderDash()`, monta o array `kpis` no início. Cada item: `{lab, val, unit, hint, c}` onde `c` é a cor da barrinha lateral.

### 5.6 O que NÃO mexer sem cuidado
- A função `migrateNames()` — só roda uma vez por instalação (controlada por flag). Se mexer no mapa, considere bumpar a flag (`cd_names_v3` etc) pra forçar nova execução.
- A função `snapshot()`/`applySnapshot()` — manter retrocompatível com o formato array antigo, senão backups antigos quebram.
- A classe `.bar` — ver seção 4.4.
- O script anti-flash no `<head>` — se remover, tema escuro pisca branco ao carregar.

---

## 6. Solução de problemas

### "Perdi meus dados depois de limpar o cache"
**Causa:** o vínculo de arquivo não estava configurado. localStorage foi apagado junto com o cache.

**Recuperação:**
- Se você tem um backup JSON anterior: menu ⋮ → Importar backup.
- Se não tem backup, o app voltou pros 1300 cartões originais — alterações posteriores foram perdidas.

**Prevenção:** vincular o arquivo de backup (selo verde no topo).

### "O selo verde virou amarelo do nada"
**Causa:** o navegador perdeu a permissão de escrita no arquivo (acontece em algumas sessões).

**Recuperação:** clique no selo amarelo → vincular de novo (geralmente o navegador lembra o arquivo).

### "Vinculei o arquivo mas ele não está salvando"
**Causa provável:** você está no Firefox ou num navegador desatualizado. A File System Access API só funciona em **Chrome/Edge atualizados**.

**Solução:** use o Chrome ou o Edge; ou use o "Exportar backup (JSON)" manualmente de tempos em tempos.

### "O modal ficou aberto e travou tudo"
Pressione **Esc** para fechar. Se não responder, clique fora do modal (na área escurecida).

### "Renomeei um vendedor mas a tabela do painel ainda mostra o nome antigo"
**Causa:** o renomeio só foi salvo no cadastro, mas o painel não foi atualizado.

**Solução:** troque de aba e volte. O painel recalcula ao abrir.

### "Tem dois 'Bruno' na lista"
Confirme o mapeamento na seção 4.5. **Bruno Almeida** e **Bruno Carvalho** ("Bruno (apelido)") são pessoas distintas, isso é intencional.

### "Preciso recomeçar do zero"
Menu ⋮ → "Restaurar dados originais". **Antes de fazer**, exporte um backup JSON pra ter como voltar atrás.

---

## 7. Publicação e demonstração

Este projeto é publicado como portfólio público. **Os dados exibidos são fictícios** — tanto clientes/obras quanto vendedores são exemplos inventados para demonstração. Em produção, o app gerencia cerca de 1300 demandas reais de uma empresa de esquadrarias de alumínio, com os dados mantidos apenas localmente (nunca versionados).

### 7.1 Estrutura do repositório
```
controle-demandas/
├── README.md
├── index.html              # o app (com dados de demonstração)
├── docs/
│   └── DOCUMENTACAO.md      # este documento
└── assets/
    └── preview.png
    └── preview2.png
    └── preview3.png
```

### 7.2 Como rodar
```bash
git clone https://github.com/JoseGavioli/controle-demandas.git
cd controle-demandas
# abra index.html no navegador (Chrome ou Edge recomendados)
```

### 7.3 Demo ao vivo
Hospedada via GitHub Pages em `https://josegavioli.github.io/controle-demandas/`. Para reproduzir: Settings → Pages → Source "Deploy from a branch" → branch `main`, pasta `/ (root)`.

### 7.4 Persistência e backup
O app salva no `localStorage` do navegador e, opcionalmente, em um arquivo `.json` real no disco via File System Access API (Chrome/Edge), com gravação a cada alteração e checkpoints automáticos. Esse mecanismo garante que os dados sobrevivam à limpeza de cache do navegador. Backups locais não são versionados (protegidos por `.gitignore`).

---

## Apêndice — Versões e marcos

| Versão | O que mudou |
|---|---|
| v1 | Kanban básico, 1300 cartões importados, modal de edição, busca/filtros, paginação do histórico |
| v2 | Painel com KPIs, tabela por vendedor, gráfico de volume mensal/anual |
| v3 | Vínculo de arquivo de backup (File System Access) + checkpoints 12h/17h |
| v4 | Normalização de nomes (migração automática única) |
| v5 | Identidade visual da marca (paleta navy/vermelho, logo embutida) |
| v6 | Cadastro de vendedores (ativar/desativar, editar, renomear com cascata) |
| v7 | Tema claro/escuro com persistência e respeito ao tema do sistema |
| v8 | Versão pública limpa para o GitHub (dados de demonstração fictícios) |
| v9 | Anonimização de vendedores na versão pública; publicação com GitHub Pages |

---

*Documento mantido por José Gavioli. Em caso de dúvida sobre uma decisão, esta documentação é a referência — o código mostra o "como", aqui está o "porquê".*
