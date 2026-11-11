# 📊 Seminário NoSQL: Apache CouchDB vs. Couchbase Server

Material acadêmico e prático completo sobre **Apache CouchDB** e **Couchbase Server**, desenvolvido sob a ótica de **Arquitetura de Dados e Sistemas Distribuídos**.

---

## 📁 Estrutura do Repositório

| Arquivo / Pasta | Descrição |
| :--- | :--- |
| **[`apresentacao.html`](./apresentacao.html)** / **[`index.html`](./index.html)** | **Apresentação Web Interativa (20 Slides)** em Dark Mode, com **Teleprompter / Roteiro de Fala integrado (Tecla `R`)**, cronômetro de 45 minutos e Quiz interativo no final. |
| **[`Apresentacao_CouchDB_vs_Couchbase.pptx`](./Apresentacao_CouchDB_vs_Couchbase.pptx)** | **Apresentação Nativa em PowerPoint (16:9)** contendo todos os slides e o roteiro de fala de cada integrante salvo nas *Notas do Apresentador*. |
| **[`workshop/docker-compose.yml`](./workshop/docker-compose.yml)** | Ambiente Docker Compose subindo `apache/couchdb:3.3` (porta `5984`) e `couchbase:community-7.6.0` (portas `8091–8096` e `11210`). |
| **[`workshop/comandos_workshop.sh`](./workshop/comandos_workshop.sh)** | Roteiro de comandos prontos (`cURL`, `couchbase-cli` e `N1QL/SQL++`) para execução rápida na demonstração ao vivo. |
| **[`workshop/crud_couchdb_demo.py`](./workshop/crud_couchdb_demo.py)** | Script em Python 3 demonstrando o fluxo CRUD real no CouchDB, controle de concorrência otimista MVCC (`_rev` / `HTTP 409 Conflict`), Mango Query (`_find`) e deleção via *Tombstone*. |
| **[`workshop/gerar_pptx.py`](./workshop/gerar_pptx.py)** | Script Python (`python-pptx`) utilizado para gerar o arquivo `.pptx`. |

---

## 👥 Distribuição da Equipe (7 Integrantes — 35 a 45 min)

| Integrante | Slides | Tema Principal | Tempo Estimado |
| :--- | :--- | :--- | :--- |
| **Integrante 1** | Slides 1, 2 e 3 | Contextualização NoSQL, História (CouchOne + Membase) e Bifurcação Filosófica | ~6 min |
| **Integrante 2** | Slides 4, 5 e 6 | Arquitetura Interna do **Apache CouchDB** (Append-Only B+Tree, MVCC `_rev`, Views/Mango e `_changes`) | ~6 min |
| **Integrante 3** | Slides 7, 8 e 9 | Arquitetura Interna do **Couchbase** (Memory-First, 1024 vBuckets, MDS, N1QL/SQL++ e GSI) | ~6 min |
| **Integrante 4** | Slides 10 e 11 | Casos de Uso, Ecossistema **PouchDB (Offline-First)**, Mercado e Trade-offs do CouchDB | ~5 min 30s |
| **Integrante 5** | Slides 12 e 13 | Casos de Uso em Escala Extrema, **XDCR**, Mercado (LinkedIn, Amadeus, PayPal) e Trade-offs do Couchbase | ~5 min 30s |
| **Integrante 6** | Slides 14 e 15 | **Quadro Comparativo Direto** (Modelo de Dados, Teorema CAP/PACELC, Consultas, Replicação e Latência) | ~5 min 30s |
| **Integrante 7** | Slides 16, 17 e 18 | **Workshop Prático ao Vivo** (Docker Compose, API REST CouchDB, N1QL Couchbase e Script Python) | ~6 min 30s |

---

## 🚀 Como Executar o Workshop Prático

### 1. Subir os Containers (CouchDB + Couchbase)
```bash
cd workshop
docker compose up -d
```

### 2. Inicializar o CouchDB e o Couchbase via CLI
```bash
# Inicializar single-node no Apache CouchDB 3.3
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

## ⚖️ Resumo da Matriz Comparativa

| Dimensão Técnica | Apache CouchDB | Couchbase Server |
| :--- | :--- | :--- |
| **Modelo de Dados** | Documentos JSON + `_attachments` em Databases planos. | Documentos JSON + KV Binário (`Bucket` $\rightarrow$ `Scope` $\rightarrow$ `Collection`). |
| **Teorema CAP / PACELC** | **AP** (PC/EC) — prioriza disponibilidade e consistência eventual. | **CP por padrão** (PC/EC) — único vBucket ativo por chave em RAM. |
| **Concorrência / ACID** | MVCC otimista explícito via árvore `_rev` no documento. | CAS otimista + Lock pessimista + Transações ACID Multi-Documento. |
| **Consultas** | Views MapReduce (JS/Erlang) + Mango Queries (`_find`). Sem `JOIN`. | **N1QL / SQL++** (`SELECT`, `ANSI JOIN`, `UNNEST`, `GROUP BY`) + FTS/Vector. |
| **Replicação** | **Couch Replication Protocol** via HTTP (`_changes` + PouchDB Offline-First). | **DCP** (intra-cluster) e **XDCR** (inter-cluster) via stream binário TCP em RAM. |
| **Performance / Latência** | Casa dos **milissegundos** (I/O append-only em disco + HTTP). | **Sub-milissegundo (microssegundos)** via cache gerenciado em memória RAM. |
