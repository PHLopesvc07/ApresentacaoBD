#!/usr/bin/env bash
# ==============================================================================
# ROTEIRO DE COMANDOS PRONTOS PARA COPIAR E COLAR NO WORKSHOP (INTEGRANTE 7)
# ==============================================================================

# ------------------------------------------------------------------------------
# PASSO 1: SUBIR OS CONTAINERS E INICIALIZAR OS CLUSTERS
# ------------------------------------------------------------------------------
docker compose up -d

# Inicializar tabelas internas do Apache CouchDB 3.3
curl -X POST http://admin:admin123@localhost:5984/_cluster_setup \
  -H "Content-Type: application/json" \
  -d '{"action": "enable_single_node", "username": "admin", "password": "admin123", "bind_address": "0.0.0.0", "port": 5984, "singlenode": true}'

# Inicializar o cluster Couchbase Server 7.6 via CLI
docker exec lab_couchbase couchbase-cli cluster-init \
  -c 127.0.0.1 \
  --cluster-username admin \
  --cluster-password admin123 \
  --services data,index,query \
  --cluster-ramsize 1024 \
  --cluster-index-ramsize 512 \
  --index-storage-setting default

# ------------------------------------------------------------------------------
# PASSO 2: PRÁTICA 1 — APACHE COUCHDB (PORTA 5984)
# ------------------------------------------------------------------------------
# 2.1 Criar o banco universidade_db
curl -X PUT http://admin:admin123@localhost:5984/universidade_db

# 2.2 Inserir documento aluno_01 (Gera revisão 1-...)
curl -X PUT http://admin:admin123@localhost:5984/universidade_db/aluno_01 \
  -H "Content-Type: application/json" \
  -d '{"nome": "Ana Carolina Silva", "curso": "Ciencia da Computacao", "semestre": 6, "cra": 8.7}'

# 2.3 Provocar conflito MVCC proposital (Tentando atualizar SEM _rev -> HTTP 409)
curl -i -X PUT http://admin:admin123@localhost:5984/universidade_db/aluno_01 \
  -H "Content-Type: application/json" \
  -d '{"nome": "Ana Carolina Silva", "curso": "Ciencia da Computacao", "semestre": 7, "cra": 9.1}'

# 2.4 Capturar _rev atual e atualizar com sucesso (Gera revisão 2-...)
REV=$(curl -s http://admin:admin123@localhost:5984/universidade_db/aluno_01 | grep -o '"_rev":"[^"]*"' | cut -d'"' -f4)
curl -X PUT http://admin:admin123@localhost:5984/universidade_db/aluno_01 \
  -H "Content-Type: application/json" \
  -d "{\"_rev\": \"$REV\", \"nome\": \"Ana Carolina Silva\", \"curso\": \"Ciencia da Computacao\", \"semestre\": 7, \"cra\": 9.1}"

# 2.5 Consulta declarativa com Mango Query (_find)
curl -X POST http://admin:admin123@localhost:5984/universidade_db/_find \
  -H "Content-Type: application/json" \
  -d '{"selector": {"cra": {"$gt": 8.5}}, "fields": ["_id", "_rev", "nome", "cra"]}'

# 2.6 Escutar Feed de Mudanças em Tempo Real (Rodar em segundo terminal)
curl -N "http://admin:admin123@localhost:5984/universidade_db/_changes?feed=continuous&include_docs=true&since=now"

# ------------------------------------------------------------------------------
# PASSO 3: PRÁTICA 2 — COUCHBASE SERVER (PORTA 8091 / 8093)
# ------------------------------------------------------------------------------
# 3.1 Criar Bucket, Scope e Collection
docker exec lab_couchbase couchbase-cli bucket-create \
  -c 127.0.0.1 -u admin -p admin123 \
  --bucket academico --bucket-type couchbase --bucket-ramsize 256

curl -s -u admin:admin123 http://localhost:8093/query/service \
  -d 'statement=CREATE SCOPE `academico`.`graduacao` IF NOT EXISTS;'

curl -s -u admin:admin123 http://localhost:8093/query/service \
  -d 'statement=CREATE COLLECTION `academico`.`graduacao`.`matriculas` IF NOT EXISTS;'

# ------------------------------------------------------------------------------
# PASSO 4: EXECUTAR O SCRIPT PYTHON CRUD COMPLETO
# ------------------------------------------------------------------------------
python crud_couchdb_demo.py
