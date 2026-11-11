# 📊 Seminário NoSQL: Apache CouchDB vs. Couchbase Server

Apresentação acadêmica interativa e laboratório prático reproduzível sobre **Apache CouchDB** e **Couchbase Server**, desenvolvidos sob a ótica de **Arquitetura de Dados, Sistemas Distribuídos e Performance em Sala de Aula (TV 50")**.

🔗 **Acesso Online (GitHub Pages — Custo Zero):** `https://phlopesvc07.github.io/ApresentacaoBD/`

---

## 🏗️ Arquitetura Setorizada do Projeto

O código foi refatorado e modularizado em camadas independentes (HTML semântico, CSS desacoplado por responsabilidade, módulos JavaScript ES6+ e scripts de laboratório) para máxima manutenibilidade e cache eficiente em CDNs gratuitas:

```text
ApresentacaoBD/
├── index.html                        # Entrypoint principal da Apresentação Web (GitHub Pages)
├── apresentacao.html                 # Espelho local da apresentação interativa
├── .nojekyll                         # Bypass de build Jekyll para deploy instantâneo no GitHub Pages
├── Apresentacao_CouchDB_vs_Couchbase_Professor.pdf # Dossiê PDF Completo para Entrega Avaliativa ao Docente
├── Apresentacao_CouchDB_vs_Couchbase.pptx  # Versão nativa PowerPoint (16:9 Dark Mode) com Notas do Orador
├── assets/
│   ├── css/
│   │   ├── variables.css             # Design Tokens, paleta de alto contraste e escala fluida TV 50"
│   │   ├── layout.css                # Stage, Cyber-Grid, CSS Containment e aceleração de hardware
│   │   └── components.css            # Cartões, simuladores, janelas de terminal, tabelas e Quiz
│   └── js/
│       ├── app.js                    # Engine de navegação, Teleprompter, Timer, Zoom e Laser (rAF)
│       └── simulators.js             # Simuladores interativos (Append-Only, Compaction, vBuckets e Quiz)
├── workshop/
│   ├── docker-compose.yml            # Orquestração dos containers Apache CouchDB 3.3 + Couchbase 7.6
│   ├── comandos_workshop.sh          # Roteiro de comandos prontos (cURL, Mango, _changes e SQL++)
│   ├── crud_couchdb_demo.py          # Script Python 3 demonstrando CRUD real + tratamento MVCC (HTTP 409)
│   ├── gerar_pptx.py                 # Gerador do PowerPoint moderno (.pptx) com Notas
│   └── gerar_pdf_professor.py        # Gerador automatizado do Dossiê Acadêmico em PDF (ReportLab)
└── README.md                         # Documentação técnica e guia operacional do seminário
```

---

## ⚡ Otimizações de Desempenho e Hospedagem Custo Zero

1. **CSS Containment (`contain: content` / `contain: strict`):**
   - Os 19 slides inativos ficam isolados da árvore de *layout/paint* do navegador. O motor gráfico renderiza apenas o slide visível no momento, garantindo **60 FPS constantes** mesmo em computadores de laboratório universitário com placa de vídeo integrada conectados a uma **TV de 50 polegadas**.
2. **Substituição de Filtros Pesados por Gradientes Pré-Calculados na GPU:**
   - Removemos filtros contínuos de `blur()` e `backdrop-filter` repetitivos nos cartões em favor de gradientes radiais compostos na GPU (`transform: translate3d(0,0,0)` e `will-change: transform`).
3. **Apontador Laser com `requestAnimationFrame` (Zero Layout Thrashing):**
   - O modo Laser (`Tecla L`) atualiza a posição do feixe sincronizado com a taxa de atualização da TV usando `translate3d` e listener `{ passive: true }`.
4. **Persistência de Sessão (`sessionStorage`) e Carregamento `defer`:**
   - Scripts carregados com atributo `defer` sem bloquear a primeira renderização (*First Contentful Paint < 50ms*). Se a página for recarregada acidentalmente durante a aula, o navegador restaura automaticamente o slide exato e a escala de zoom escolhida para a TV.
5. **Pronto para Hospedagem Estática 100% Gratuita:**
   - Compatível nativamente com **GitHub Pages** (arquivo `.nojekyll` incluso), **Cloudflare Pages**, **Vercel** ou **Netlify**, além de funcionar **100% offline** abrindo o `index.html` localmente.

---

## 🎮 Controles de Palco, Simuladores e Atalhos (TV de 50 Polegadas)

| Tecla / Botão | Função na Apresentação |
| :--- | :--- |
| **`→` / `Espaço`** | Avança para o próximo slide |
| **`←`** | Retorna ao slide anterior |
| **`+` / `-` / `0`** | **Zoom Fluido TV 50":** Aumenta, diminui ou reseta (`112%`) o tamanho global das fontes na hora |
| **`R`** | **Teleprompter Flutuante:** Abre/fecha a gaveta lateral com o **Roteiro de Fala** sem encolher o slide na TV |
| **`G`** | **Mapa Geral de Slides:** Exibe grade com todos os 20 slides e o integrante responsável por cada um |
| **`L`** | **Apontador Laser Virtual:** Transforma o cursor em um ponto de laser vermelho brilhante na TV |
| **`F`** | Alterna o modo **Tela Cheia (Fullscreen)** |
| **`⏱ 00:00`** | Clique no relógio do topo para iniciar/pausar o cronômetro geral da apresentação |

### 🧪 Simuladores Arquiteturais Embutidos nos Slides:
- **Slide 4 (Simulador Append-Only & Compaction do CouchDB):** Botões interativos que adicionam novos blocos ao final do arquivo `.couch`, marcam blocos antigos como órfãos e executam o *Compaction* ao vivo.
- **Slide 7 (Simulador CRC32 $\rightarrow$ vBucket no Couchbase):** Calcula o hash da chave, identifica o `vBucket #` (de `0` a `1023`) e demonstra o salto único (*Single-Hop TCP*) direto para a RAM do nó ativo.
- **Slides 17 e 18 (Terminais com Simulação de Output):** Botões que exibem instantaneamente o retorno do `HTTP 409 Conflict`, o stream contínuo `_changes` e o plano de execução `N1QL (SQL++)` na tela.
- **Slides 19 e 20 (Quiz Universitário Interativo):** Clique em qualquer alternativa (`A` a `E`) para validar a resposta e exibir o gabarito comentado para a turma.

---

## 👥 Distribuição da Equipe (7 Integrantes — 35 a 45 min)

| Integrante | Slides | Foco Técnico | Tempo Estimado |
| :--- | :--- | :--- | :--- |
| **Integrante 1** | Slides 1, 2 e 3 | Contextualização NoSQL, História (CouchOne + Membase) e Bifurcação Filosófica | ~6 min |
| **Integrante 2** | Slides 4, 5 e 6 | Arquitetura Interna do **Apache CouchDB** (Append-Only B+Tree, MVCC `_rev`, Views/Mango e `_changes`) | ~6 min |
| **Integrante 3** | Slides 7, 8 e 9 | Arquitetura Interna do **Couchbase** (Memory-First, 1024 vBuckets, MDS, N1QL/SQL++ e GSI) | ~6 min |
| **Integrante 4** | Slides 10 e 11 | Casos de Uso, Ecossistema **PouchDB (Offline-First)**, Mercado e Trade-offs do CouchDB | ~5 min 30s |
| **Integrante 5** | Slides 12 e 13 | Casos de Uso em Escala Extrema, **XDCR**, Mercado (LinkedIn, Amadeus, PayPal) e Trade-offs do Couchbase | ~5 min 30s |
| **Integrante 6** | Slides 14 e 15 | **Quadro Comparativo Direto** (Modelo de Dados, Teorema CAP/PACELC, Consultas, Replicação e Latência) | ~5 min 30s |
| **Integrante 7** | Slides 16, 17 e 18 | **Workshop Prático ao Vivo** (Docker Compose, API REST CouchDB, N1QL Couchbase e Script Python) | ~6 min 30s |

---

## 🚀 Como Executar o Workshop Prático Localmente

### 1. Subir os Containers (CouchDB 3.3 + Couchbase 7.6)
```bash
cd workshop
docker compose up -d
```

### 2. Inicializar os Serviços via CLI
```bash
# Inicializar tabelas de sistema do Apache CouchDB 3.3
curl -X POST http://admin:admin123@localhost:5984/_cluster_setup \
  -H "Content-Type: application/json" \
  -d '{"action": "enable_single_node", "username": "admin", "password": "admin123", "bind_address": "0.0.0.0", "port": 5984, "singlenode": true}'

# Inicializar o cluster Couchbase Server 7.6 (Serviços: Data, Index e Query)
docker exec lab_couchbase couchbase-cli cluster-init \
  -c 127.0.0.1 \
  --cluster-username admin \
  --cluster-password admin123 \
  --services data,index,query \
  --cluster-ramsize 1024 \
  --cluster-index-ramsize 512 \
  --index-storage-setting default
```

- **Interface Web do CouchDB (Fauxton):** `http://localhost:5984/_utils` (`admin` / `admin123`)
- **Interface Web do Couchbase Console:** `http://localhost:8091` (`admin` / `admin123`)

### 3. Executar o Script Python de Demonstração CRUD + MVCC
```bash
pip install requests
python workshop/crud_couchdb_demo.py
```

---

## ⚖️ Síntese da Matriz Comparativa

| Dimensão Técnica | Apache CouchDB | Couchbase Server |
| :--- | :--- | :--- |
| **Modelo de Dados** | Documentos JSON + `_attachments` em Databases planos. | Documentos JSON + KV Binário (`Bucket` $\rightarrow$ `Scope` $\rightarrow$ `Collection`). |
| **Teorema CAP / PACELC** | **AP** (PC/EC) — prioriza disponibilidade e consistência eventual. | **CP por padrão** (PC/EC) — único vBucket ativo por chave em RAM. |
| **Concorrência / ACID** | MVCC otimista explícito via árvore `_rev` no documento. | CAS otimista + Lock pessimista + Transações ACID Multi-Documento. |
| **Consultas** | Views MapReduce (JS/Erlang) + Mango Queries (`_find`). Sem `JOIN`. | **N1QL / SQL++** (`SELECT`, `ANSI JOIN`, `UNNEST`, `GROUP BY`) + FTS/Vector. |
| **Replicação** | **Couch Replication Protocol** via HTTP (`_changes` + PouchDB Offline-First). | **DCP** (intra-cluster) e **XDCR** (inter-cluster) via stream binário TCP em RAM. |
| **Performance / Latência** | Casa dos **milissegundos** (I/O append-only em disco + HTTP). | **Sub-milissegundo (microssegundos)** via cache gerenciado em memória RAM. |
