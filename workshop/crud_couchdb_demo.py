"""
Workshop Universitário: Demonstração Prática de CRUD + Controle de Concorrência MVCC
Banco de Dados: Apache CouchDB 3.3
Requisito: pip install requests
"""

import requests
import json

# Configurações de conexão HTTP REST nativa
BASE_URL = "http://localhost:5984"
DB_NAME = "universidade_db"
AUTH = ("admin", "admin123")
HEADERS = {"Content-Type": "application/json"}


def garantir_banco():
    """Garante que o banco de dados exista no CouchDB."""
    resp = requests.put(f"{BASE_URL}/{DB_NAME}", auth=AUTH)
    if resp.status_code in (201, 412):  # 201 Created ou 412 Already Exists
        print(f"[OK] Banco '{DB_NAME}' pronto para uso.")
    else:
        resp.raise_for_status()


def create_ou_update_documento(doc_id: str, payload: dict) -> str:
    """CREATE / UPDATE: Insere ou atualiza um documento JSON retornando o novo _rev."""
    url = f"{BASE_URL}/{DB_NAME}/{doc_id}"
    resp = requests.put(url, auth=AUTH, headers=HEADERS, data=json.dumps(payload))

    if resp.status_code == 409:
        raise RuntimeError(
            f"[CONFLITO MVCC - HTTP 409] O documento '{doc_id}' possui uma revisão mais recente no banco! "
            "Envie o campo '_rev' atualizado."
        )

    resp.raise_for_status()
    dados = resp.json()
    print(f"[WRITE OK] ID: {dados['id']} | Nova Revisão (_rev): {dados['rev']}")
    return dados["rev"]


def read_documento(doc_id: str) -> dict:
    """READ: Busca um documento pela chave primária (_id) junto com seu _rev atual."""
    url = f"{BASE_URL}/{DB_NAME}/{doc_id}"
    resp = requests.get(url, auth=AUTH)
    resp.raise_for_status()
    doc = resp.json()
    print(f"[READ OK] Documento lido: {json.dumps(doc, ensure_ascii=False)}")
    return doc


def query_mango(curso: str, cra_minimo: float):
    """QUERY: Executa uma consulta declarativa via endpoint Mango (_find)."""
    url = f"{BASE_URL}/{DB_NAME}/_find"
    filtro = {
        "selector": {
            "curso": {"$eq": curso},
            "cra": {"$gte": cra_minimo}
        },
        "fields": ["_id", "_rev", "nome", "curso", "cra"]
    }
    resp = requests.post(url, auth=AUTH, headers=HEADERS, data=json.dumps(filtro))
    resp.raise_for_status()
    resultados = resp.json().get("docs", [])
    print(f"[MANGO QUERY] Encontrados {len(resultados)} aluno(s) de '{curso}' com CRA >= {cra_minimo}:")
    for item in resultados:
        print(f"   -> {item['_id']}: {item['nome']} (CRA: {item['cra']})")


def delete_documento(doc_id: str, rev: str):
    """DELETE: Cria uma revisão tombstone (_deleted: true) exigindo o _rev atual."""
    url = f"{BASE_URL}/{DB_NAME}/{doc_id}"
    resp = requests.delete(url, auth=AUTH, params={"rev": rev})
    resp.raise_for_status()
    dados = resp.json()
    print(f"[DELETE OK] Tombstone gravado para '{doc_id}' | Revisão final: {dados['rev']}")


if __name__ == "__main__":
    print("=== INICIANDO FLUXO CRUD COM CONTROLE MVCC NO APACHE COUCHDB ===\n")
    garantir_banco()

    doc_id = "aluno_demo_python"

    # Limpeza prévia caso o documento já exista de um teste anterior
    check = requests.get(f"{BASE_URL}/{DB_NAME}/{doc_id}", auth=AUTH)
    if check.status_code == 200:
        delete_documento(doc_id, check.json()["_rev"])

    # 1. CREATE: Inserir documento inicial (Gera revisão 1-...)
    novo_aluno = {
        "nome": "Lucas Martins",
        "curso": "Ciencia da Computacao",
        "semestre": 5,
        "cra": 8.4
    }
    rev_v1 = create_ou_update_documento(doc_id, novo_aluno)

    # 2. READ: Ler o documento recém-criado
    doc_atual = read_documento(doc_id)

    # 3. DEMONSTRAÇÃO DE CONFLITO MVCC: Tentar atualizar sem informar o _rev
    print("\n--- Simulando escrita concorrente sem enviar o '_rev' ---")
    try:
        create_ou_update_documento(doc_id, {"nome": "Lucas Martins", "cra": 9.0})
    except RuntimeError as erro:
        print(erro)

    # 4. UPDATE CORRETO: Passando o _rev capturado na leitura (Gera revisão 2-...)
    print("\n--- Atualizando corretamente enviando o '_rev' atual ---")
    doc_atual["cra"] = 9.3
    doc_atual["semestre"] = 6
    rev_v2 = create_ou_update_documento(doc_id, doc_atual)

    # 5. QUERY: Consultar via Mango Query (_find)
    print("\n--- Consultando documentos via Mango Query (_find) ---")
    query_mango(curso="Ciencia da Computacao", cra_minimo=8.5)

    # 6. DELETE: Remover o documento passando a revisão v2
    print("\n--- Removendo o documento (Append-Only Tombstone) ---")
    delete_documento(doc_id, rev_v2)
