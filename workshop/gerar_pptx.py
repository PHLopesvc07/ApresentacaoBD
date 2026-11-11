"""
Gerador da Apresentação PowerPoint (.pptx) com Design Moderno, Dinâmico e Profissional
Tema: Widescreen 16:9, Glassmorphism Dark Mode, Gradientes, Badges e Tipografia Otimizada
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Paleta de Cores Moderna (Cyber/Executive Dark Mode)
BG_COLOR = RGBColor(6, 10, 20)          # #060a14 - Fundo escuro profundo
CARD_BG = RGBColor(16, 24, 43)          # #10182b - Card principal
CARD_ALT_BG = RGBColor(22, 33, 58)      # #16213a - Card secundário
TOPBAR_BG = RGBColor(11, 17, 31)        # #0b111f - Barra superior
BORDER_DEFAULT = RGBColor(38, 52, 82)   # #263452 - Borda sutil glass
BORDER_ORANGE = RGBColor(255, 107, 43)  # #ff6b2b - CouchDB Laranja vibrante
BORDER_BLUE = RGBColor(0, 210, 255)     # #00d2ff - Couchbase Ciano neon
BORDER_GREEN = RGBColor(16, 185, 129)   # #10b981 - Verde esmeralda
BORDER_PURPLE = RGBColor(192, 132, 252) # #c084fc - Roxo neon
TEXT_WHITE = RGBColor(255, 255, 255)    # Branco puro
TEXT_MUTED = RGBColor(176, 190, 210)    # Cinza claro legível
ACCENT_YELLOW = RGBColor(253, 224, 71)  # Destaque #fde047
ACCENT_CYAN = RGBColor(125, 211, 252)   # Código/Chaves

SLIDES_DATA = [
    {
        "meta": "MÓDULO 1 • ORIGEM E FUNDAMENTOS ARQUITETURAIS",
        "title": "Apache CouchDB vs. Couchbase Server",
        "subtitle": "Arquiteturas Documentais: Da Sincronização Offline-First à Performance Sub-Milissegundo em Memória",
        "speaker": "Integrante 1 (Tempo: 2 min | Acumulado: 02:00)",
        "theme": "mixed",
        "card1": {
            "title": "🌐 A Crise do RDBMS Clássico",
            "border": BORDER_ORANGE,
            "items": [
                "Limites físicos da escalabilidade vertical (Scale-Up) na Web 2.0 e Mobile.",
                "Impedância Objeto-Relacional: Custo extremo de JOINs em dezenas de tabelas.",
                "Bloqueios pessimistas (2PL) inviáveis em clusters geograficamente distribuídos."
            ]
        },
        "card2": {
            "title": "⚡ O Paradigma NoSQL Documental",
            "border": BORDER_BLUE,
            "items": [
                "Documentos auto-contidos em JSON com objetos e arrays aninhados.",
                "Localidade de Dados: Leitura de um agregado completo em um único salto de I/O.",
                "Schema-on-read: Flexibilidade sem bloqueios de migração ALTER TABLE."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 1 (2 minutos):\n\n"
            "Bom dia/Boa tarde a todos. Hoje nossa equipe vai apresentar uma análise arquitetural profunda sobre dois bancos de dados NoSQL orientados a documentos que frequentemente causam confusão pelo nome, mas que resolveram problemas completamente distintos na engenharia de software: o Apache CouchDB e o Couchbase.\n\n"
            "Quando olhamos para as limitações dos bancos relacionais tradicionais em sistemas distribuídos de larga escala, esbarramos em dois gargalos clássicos: a impedância objeto-relacional e o custo de manter transações distribuídas com bloqueio pessimista. Tanto o CouchDB quanto o Couchbase adotam o JSON como unidade básica de armazenamento, mas enquanto um nasceu para dominar a web descentralizada e redes instáveis, o outro foi projetado para entregar latência sub-milissegundo em clusters massivos orientados a memória."
        )
    },
    {
        "meta": "SLIDE 2 • GENEALOGIA HISTÓRICA E FUSÃO",
        "title": "Origem Histórica: A Fusão de 2011 e a Grande Divergência",
        "subtitle": "Como CouchOne + Membase formaram a Couchbase Inc. enquanto o Apache CouchDB seguiu independente",
        "speaker": "Integrante 1 (Tempo: 2 min | Acumulado: 04:00)",
        "theme": "mixed",
        "card1": {
            "title": "🪶 O Caminho do Apache CouchDB (ASF)",
            "border": BORDER_ORANGE,
            "items": [
                "2005: Damien Katz (ex-IBM Lotus Notes) cria o CouchDB em Erlang/OTP.",
                "2008: Torna-se projeto Top-Level oficial da Apache Software Foundation.",
                "2009: Fundação da CouchOne para suporte corporativo ao CouchDB.",
                "Fiel à arquitetura: Armazenamento Append-Only e protocolo HTTP/REST nativo."
            ]
        },
        "card2": {
            "title": "⚡ A Criação do Couchbase Server",
            "border": BORDER_BLUE,
            "items": [
                "2010: Surgimento do Membase (persistência + replicação para o Memcached).",
                "Fevereiro 2011: FUSÃO CouchOne + Membase = Couchbase Inc.",
                "Divergência: Substituição do motor do CouchDB pelo motor em RAM do Membase.",
                "Adição de particionamento por vBuckets, N1QL (SQL++) e índices GSI."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 1 (2 minutos):\n\n"
            "Para entender por que esses dois bancos têm 'Couch' no nome, precisamos voltar a 2005. Damien Katz criou o CouchDB inspirado em bancos distribuídos replicados. Em paralelo, a equipe por trás do Memcached criou o Membase. Em 2011, houve a fusão oficial formando a Couchbase Inc.\n\n"
            "O ponto crucial é que o Couchbase Server não é um CouchDB renomeado. A engenharia do Couchbase uniu o motor de chave-valor em memória do Membase com indexação JSON. Enquanto isso, a comunidade aberta manteve o Apache CouchDB evoluindo de forma independente sob a Apache Foundation."
        )
    },
    {
        "meta": "SLIDE 3 • BIFURCAÇÃO FILOSÓFICA FUNDAMENTAL",
        "title": "Bifurcação Filosófica: Web Nativa vs. Memory-First",
        "subtitle": "Decisões opostas de protocolos, hierarquia de memória e tolerância a partições",
        "speaker": "Integrante 1 (Tempo: 2 min | Acumulado: 06:00)",
        "theme": "mixed",
        "card1": {
            "title": "🟧 Apache CouchDB: Sync & Web-First",
            "border": BORDER_ORANGE,
            "items": [
                "Protocolo 100% HTTP/REST nativo (porta 5984) acessível sem drivers especiais.",
                "Caminho Disk-First: Append-Only B+Tree imune a quedas de energia.",
                "Concorrência: MVCC explícito no documento via campo _rev.",
                "Foco: Tolerância a partições (AP), replicação Multi-Master e Offline-First."
            ]
        },
        "card2": {
            "title": "🟦 Couchbase Server: Memory-First & Scale",
            "border": BORDER_BLUE,
            "items": [
                "Protocolo binário sobre TCP (porta 11210) com SDK inteligente ciente do mapa do cluster.",
                "Caminho Memory-First: Leituras e escritas servidas diretamente do cache em RAM.",
                "Consistência: vBucket único ativo por chave (CP por padrão) + CAS e transações ACID.",
                "Foco: Vazão extrema, latência em microssegundos e consultas SQL++ (N1QL)."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 1 (2 minutos):\n\n"
            "Essa separação gerou duas filosofias de engenharia opostas. O Apache CouchDB fala HTTP nativamente, grava em disco append-only e expõe o campo _rev para garantir sincronização entre servidores desconectados.\n\n"
            "Já o Couchbase coloca um cache gerenciado em RAM na frente do disco, confirmando operações em microssegundos e roteando via vBuckets. Agora passo a palavra para o Integrante 2, que vai dissecar o motor interno do Apache CouchDB."
        )
    },
    {
        "meta": "SLIDE 4 • ARQUITETURA INTERNA APACHE COUCHDB",
        "title": "CouchDB: Motor Append-Only e Indexação B+Tree",
        "subtitle": "Imutabilidade em disco, concorrência lock-free em Erlang e imunidade a corrupção",
        "speaker": "Integrante 2 (Tempo: 2 min | Acumulado: 08:00)",
        "theme": "couchdb",
        "card1": {
            "title": "⚙️ Motor Append-Only (Tail-Append)",
            "border": BORDER_ORANGE,
            "items": [
                "Construído em Erlang/OTP: Processos leves isolados e tolerância a falhas (Let it crash).",
                "Zero In-Place Updates: Blocos de disco jamais são sobrescritos.",
                "Novas revisões, nós alterados da B+Tree e o Tail Header vão para o final do arquivo.",
                "Tail Header no fim do arquivo aponta para a nova raiz da árvore B+ após cada commit."
            ]
        },
        "card2": {
            "title": "🛡️ Garantias e Processo de Compaction",
            "border": BORDER_GREEN,
            "items": [
                "Leituras Lock-Free: Leitores navegam snapshots imutáveis sem bloquear escritores.",
                "Imunidade a corrupção: Em queda de energia, varre de trás para frente até o último header íntegro.",
                "Compaction em background: Reescreve o arquivo mantendo apenas a revisão ativa mais recente.",
                "Trade-off: Crescimento contínuo de disco até a execução do Compaction."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 2 (2 minutos):\n\n"
            "Obrigado. Vamos mergulhar na engenharia do CouchDB. Escrito em Erlang, seu núcleo é a B+Tree Append-Only. Ele jamais altera bytes já gravados; qualquer modificação grava novos dados no final do arquivo junto a um novo Tail Header.\n\n"
            "Isso traz duas vantagens monumentais: leituras nunca bloqueiam escritas, e o banco é imune a corrupções por queda de energia. O trade-off é que arquivos crescem rápido, exigindo o processo periódico de Compaction."
        )
    },
    {
        "meta": "SLIDE 5 • CONTROLE DE CONCORRÊNCIA NO COUCHDB",
        "title": "CouchDB: MVCC Explícito (_rev) e Árvore de Revisões",
        "subtitle": "Como cada documento funciona como um mini-repositório Git para resolver conflitos distribuídos",
        "speaker": "Integrante 2 (Tempo: 2 min | Acumulado: 10:00)",
        "theme": "couchdb",
        "card1": {
            "title": "🔐 Ciclo Optimistic Locking (_id + _rev)",
            "border": BORDER_ORANGE,
            "items": [
                "Identificador de revisão: <geração>-<hash_md5> (ex: 2-7051cbe5...).",
                "Atualizações exigem que o cliente envie o _rev atual na requisição HTTP PUT.",
                "Conflito imediato: Se o _rev for obsoleto, retorna HTTP 409 Conflict.",
                "Deleções gravam um Tombstone (\"_deleted\": true) preservando a linha histórica."
            ]
        },
        "card2": {
            "title": "🌿 Conflitos na Replicação Master-Master",
            "border": BORDER_PURPLE,
            "items": [
                "Edições simultâneas em nós offline criam ramificações na árvore de revisões.",
                "Eleição determinística: Todos os nós elegem o mesmo vencedor comparando hashes (sem NTP!).",
                "Revisões perdedoras são mantidas em _conflicts até que a aplicação faça o merge.",
                "Zero perda silenciosa de dados em partições de rede."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 2 (2 minutos):\n\n"
            "O CouchDB garante consistência através de MVCC no nível de documento. Cada documento possui um histórico de revisões. Para alterar um registro, você deve enviar o _rev atual. Se alguém atualizou antes, o CouchDB responde com HTTP 409 Conflict.\n\n"
            "Em replicação distribuída entre nós desconectados, se houver escritas concorrentes, o CouchDB mantém ambas as revisões e usa um algoritmo determinístico baseado no hash para eleger a versão vencedora em todos os servidores sem depender de relógios físicos."
        )
    },
    {
        "meta": "SLIDE 6 • CONSULTAS E REPLICAÇÃO NO COUCHDB",
        "title": "CouchDB: MapReduce, Mango Query e _changes Feed",
        "subtitle": "Indexação incremental em B-Tree e o protocolo de replicação contínua em tempo real",
        "speaker": "Integrante 2 (Tempo: 2 min | Acumulado: 12:00)",
        "theme": "couchdb",
        "card1": {
            "title": "🗺️ Consultas: MapReduce & Mango (_find)",
            "border": BORDER_ORANGE,
            "items": [
                "Design Documents (_design/): Funções map(doc) em JavaScript emitindo chaves para B-Trees.",
                "Reduções pré-computadas (_count, _sum) armazenadas nos nós internos da B-Tree.",
                "Mango Query (POST /_find): Sintaxe declarativa em JSON introduzida no CouchDB 2.0+.",
                "Suporte a índices secundários declarativos via POST /_index sem escrever código manual."
            ]
        },
        "card2": {
            "title": "📡 O Endpoint _changes e Replicação",
            "border": BORDER_GREEN,
            "items": [
                "Log de eventos sequencial (update_seq) exposto em tempo real via HTTP Stream.",
                "Couch Replication Protocol: Origem lê _changes e pergunta deltas via _revs_diff.",
                "Transferência de documentos em lote via _bulk_docs com reconexão automática.",
                "Compatibilidade nativa com PouchDB no navegador e dispositivos móveis."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 2 (2 minutos):\n\n"
            "Para buscas, o CouchDB oferece Views MapReduce em JavaScript e o moderno Mango Query, que usa filtros declarativos em JSON. Mas o verdadeiro motor do CouchDB é o endpoint _changes.\n\n"
            "Como o armazenamento é append-only, toda mutação ganha um número de sequência. O protocolo de replicação lê esse feed, identifica diferenças e transfere apenas os deltas. Passo agora ao Integrante 3 para explicar a arquitetura do Couchbase."
        )
    },
    {
        "meta": "SLIDE 7 • ARQUITETURA INTERNA COUCHBASE SERVER",
        "title": "Couchbase: Memory-First e Particionamento por vBuckets",
        "subtitle": "Cache gerenciado integrado, protocolo binário DCP e roteamento direto Single-Hop",
        "speaker": "Integrante 3 (Tempo: 2 min | Acumulado: 14:00)",
        "theme": "couchbase",
        "card1": {
            "title": "🚀 Caminho de Dados Memory-First (KV)",
            "border": BORDER_BLUE,
            "items": [
                "Leituras e escritas passam obrigatoriamente pela memória RAM gerenciada.",
                "Confirmação de escrita para a aplicação em microssegundos (sub-milissegundo).",
                "Protocolo interno DCP: Stream binário em memória para filas assíncronas.",
                "Replication Queue alimenta réplicas em RAM; Persistence Queue grava em disco NVMe."
            ]
        },
        "card2": {
            "title": "🧩 Particionamento Fixo: 1024 vBuckets",
            "border": BORDER_GREEN,
            "items": [
                "Todo bucket é dividido estaticamente em exatamente 1024 vBuckets.",
                "Roteamento determinístico no SDK: vBucket_ID = CRC32(Key) % 1024.",
                "SDK mantém o Cluster Map em RAM e conecta direto no nó ativo daquele vBucket.",
                "Zero nós proxies ou gargalos intermediários (Roteamento Single-Hop)."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 3 (2 minutos):\n\n"
            "Obrigado. No Couchbase mudamos completamente de paradigma: é uma arquitetura Memory-First. As escritas vão para a RAM gerenciada pelo Data Service e retornam confirmação imediata em frações de milissegundo, enquanto o protocolo interno DCP propaga os dados para réplicas e disco.\n\n"
            "Para distribuir a carga sem proxies que adicionam latência, o Couchbase divide cada bucket em 1024 vBuckets. O SDK calcula o hash CRC32 da chave e fala diretamente com o nó que possui o vBucket ativo em um único salto TCP."
        )
    },
    {
        "meta": "SLIDE 8 • MULTI-DIMENSIONAL SCALING (MDS)",
        "title": "Couchbase: Arquitetura Multi-Service e Isolamento MDS",
        "subtitle": "Eliminando o efeito 'Noisy Neighbor' ao desacoplar CPU, RAM e I/O em nós especializados",
        "speaker": "Integrante 3 (Tempo: 2 min | Acumulado: 16:00)",
        "theme": "couchbase",
        "card1": {
            "title": "❌ O Problema dos Clusters Homogêneos",
            "border": BORDER_ORANGE,
            "items": [
                "Bancos NoSQL clássicos executam armazenamento, índices e queries no mesmo nó.",
                "Efeito Vizinho Barulhento: Uma consulta analítica pesada consome toda a CPU.",
                "Queda drástica de throughput nas operações de leitura e escrita operacionais."
            ]
        },
        "card2": {
            "title": "✔ Solução MDS do Couchbase (6 Serviços)",
            "border": BORDER_BLUE,
            "items": [
                "Data Service (KV): Nós com alta RAM para operações por chave sub-ms.",
                "Index Service (GSI): Nós com RAM/NVMe mantendo índices globais secundários.",
                "Query Service (SQL++): Nós com múltiplos núcleos de CPU para parsing e joins.",
                "Search (FTS e Vetores IA), Analytics (MPP Colunar) e Eventing (Serverless JS)."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 3 (2 minutos):\n\n"
            "Outro grande diferencial do Couchbase é o Multi-Dimensional Scaling (MDS). Em clusters homogêneos tradicionais, todos os nós fazem tudo, o que causa interferência entre queries pesadas e leituras simples.\n\n"
            "Com o MDS, o Couchbase desacopla o banco em serviços independentes: Data para memória, Query para processamento de CPU e Index para SSDs rápidos. Mutações fluem entre os nós via protocolo DCP sem competir por recursos."
        )
    },
    {
        "meta": "SLIDE 9 • N1QL (SQL++) E ÍNDICES GSI",
        "title": "Couchbase: Hierarquia Lógica, SQL++ (N1QL) e GSI",
        "subtitle": "A expressividade do ANSI SQL adaptada nativamente a documentos JSON aninhados",
        "speaker": "Integrante 3 (Tempo: 2 min | Acumulado: 18:00)",
        "theme": "couchbase",
        "card1": {
            "title": "🗂️ Mapeamento Relacional e N1QL",
            "border": BORDER_BLUE,
            "items": [
                "Hierarquia em 3 níveis (v7+): Bucket (DB) ➔ Scope (Schema) ➔ Collection (Tabela).",
                "N1QL / SQL++: Suporte completo a ANSI JOINs entre documentos distintos.",
                "Operador UNNEST: Transforma arrays internos do JSON em linhas relacionais.",
                "Suporte a transações ACID multi-documento distribuídas no cluster."
            ]
        },
        "card2": {
            "title": "⚡ Global Secondary Indexes (GSI)",
            "border": BORDER_GREEN,
            "items": [
                "Índices globais concentrados ou particionados nos nós de Index Service.",
                "Evita Scatter-Gather: O Query Service consulta o GSI e vai direto no vBucket exato.",
                "Memory-Optimized Indexes (MOI): Índices mantidos 100% em RAM usando Skip-Lists.",
                "Covering Indexes respondem queries complexas em menos de 2 milissegundos."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 3 (2 minutos):\n\n"
            "A partir da versão 7, o Couchbase adotou a hierarquia Bucket, Scope e Collection, facilitando a transição de quem vem do modelo relacional. E com o N1QL (SQL++), você escreve ANSI SQL padrão com JOINs, agregações e o operador UNNEST para desmembrar arrays JSON.\n\n"
            "Para acelerar o SQL++, o Couchbase usa GSI (Global Secondary Indexes), evitando varrer todos os nós do cluster. Passo a palavra agora ao Integrante 4 para discutir os casos de mercado do CouchDB."
        )
    },
    {
        "meta": "SLIDE 10 • CASOS DE USO DO APACHE COUCHDB",
        "title": "Onde o Apache CouchDB Brilha: Offline-First e Edge",
        "subtitle": "Sincronização bidirecional resiliente com PouchDB em ambientes sem conectividade",
        "speaker": "Integrante 4 (Tempo: 2m45s | Acumulado: 20:45)",
        "theme": "couchdb",
        "card1": {
            "title": "📱 Ecossistema CouchDB + PouchDB",
            "border": BORDER_ORANGE,
            "items": [
                "PouchDB executa no navegador (IndexedDB) ou mobile/Electron (SQLite).",
                "Aplicações funcionam 100% desconectadas com zero latência de rede.",
                "Sincronização bidirecional em uma linha: localDB.sync(remoteCouchDB).",
                "Retomada automática após quedas de conexão sem perda de transações."
            ]
        },
        "card2": {
            "title": "🛰️ Cenários Ideais de Engenharia",
            "border": BORDER_GREEN,
            "items": [
                "Operações remotas de campo: Mineração, plataformas de petróleo, aviação e saúde rural.",
                "Dispositivos IoT e Edge Computing com baixa memória (64 MB a 256 MB de RAM).",
                "Repositórios descentralizados em filiais autônomas com sincronização via satélite.",
                "Backends leves onde a infraestrutura não justifica clusters complexos em memória."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 4 (2 minutos e 45 segundos):\n\n"
            "Obrigado. Quando escolher o Apache CouchDB? A resposta é clara: quando o foco for Offline-First e Sincronização na Borda. Pareado ao PouchDB, ele permite que aplicações web e mobile leiam e gravem localmente no aparelho.\n\n"
            "Imagine técnicos de campo em plataformas de petróleo ou equipes de socorro em áreas sem internet. Os dados são salvos localmente e, quando surge sinal por poucos segundos, sincronizam automaticamente com a nuvem via HTTP."
        )
    },
    {
        "meta": "SLIDE 11 • MERCADO E LIMITAÇÕES DO COUCHDB",
        "title": "Apache CouchDB no Mundo Real e Seus Trade-offs",
        "subtitle": "De contratos inteligentes (Hyperledger Fabric) aos gargalos de escritas massivas",
        "speaker": "Integrante 4 (Tempo: 2m45s | Acumulado: 23:30)",
        "theme": "couchdb",
        "card1": {
            "title": "🏆 Cases Reais no Mercado",
            "border": BORDER_GREEN,
            "items": [
                "Hyperledger Fabric (Linux Foundation): State Database oficial para Smart Contracts em JSON.",
                "npm Registry: Histórico de distribuição e espelhamento global de pacotes JavaScript.",
                "eHealth Africa / Médicos Sem Fronteiras: Rastreamento em surtos de Ebola com tablets offline.",
                "IBM Cloudant: Serviço corporativo totalmente gerenciado baseado no Apache CouchDB."
            ]
        },
        "card2": {
            "title": "⚠️ Trade-offs (Quando NÃO Usar)",
            "border": BORDER_ORANGE,
            "items": [
                "Contadores e telemetria atualizados centenas de vezes por segundo na mesma chave.",
                "Crescimento exponencial de disco devido a revisões _rev pendentes de Compaction.",
                "Inexistência de JOINs no servidor e transações ACID multi-documento.",
                "Latência de dezenas de milissegundos limitada por I/O append-only e HTTP."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 4 (2 minutos e 45 segundos):\n\n"
            "Casos emblemáticos incluem o Hyperledger Fabric, que usa CouchDB para guardar o estado mundial de contratos inteligentes em JSON auditável, e o npm registry. No entanto, é fundamental conhecer os trade-offs.\n\n"
            "Não utilize o CouchDB para sistemas de alta frequência que atualizam o mesmo registro mil vezes por segundo, pois a árvore de revisões e o disco inflarão rapidamente. Agora passo a palavra ao Integrante 5 para detalhar o Couchbase."
        )
    },
    {
        "meta": "SLIDE 12 • CASOS DE USO DO COUCHBASE SERVER",
        "title": "Onde o Couchbase Brilha: Escala Extrema, Sub-ms e XDCR",
        "subtitle": "Unificando Cache Distribuído + Banco Documental com replicação entre datacenters",
        "speaker": "Integrante 5 (Tempo: 2m45s | Acumulado: 26:15)",
        "theme": "couchbase",
        "card1": {
            "title": "⚡ Eliminação da Camada Dupla (Cache + DB)",
            "border": BORDER_BLUE,
            "items": [
                "Substitui a arquitetura híbrida RDBMS/MongoDB + Redis/Memcached por um só cluster.",
                "Elimina problemas de sincronização dupla e invalidação de cache (Cache Invalidation).",
                "Ideal para sessões de usuários globais, perfis 360°, carrinhos de e-commerce e jogos online.",
                "Processamento contínuo de centenas de milhares de operações por segundo."
            ]
        },
        "card2": {
            "title": "🌍 XDCR (Cross Data Center Replication)",
            "border": BORDER_PURPLE,
            "items": [
                "Replicação assíncrona memória-a-memória via TCP direto entre múltiplos datacenters.",
                "Topologias Active-Active bidirecionais entre nuvens distintas (AWS ⇄ GCP ⇄ Azure).",
                "Resolução de conflitos automática por Sequence Number ou HLC (Hybrid Logical Clocks).",
                "Filtragem granular para replicar apenas documentos e coleções específicas."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 5 (2 minutos e 45 segundos):\n\n"
            "Obrigado. Quando entramos no Couchbase, falamos de missão crítica em larga escala. Em muitas empresas, coloca-se um banco relacional no fundo e um Redis na frente. O Couchbase resolve isso porque ele já é o cache e o banco persistente ao mesmo tempo.\n\n"
            "Além disso, para atuação global, o XDCR replica mutações diretamente da memória de um datacenter no Brasil para outro nos Estados Unidos usando streams TCP binários e relógios lógicos híbridos."
        )
    },
    {
        "meta": "SLIDE 13 • MERCADO E LIMITAÇÕES DO COUCHBASE",
        "title": "Couchbase em Escala Global e Seus Trade-offs",
        "subtitle": "Cases de milhões de operações por segundo e custos operacionais de memória",
        "speaker": "Integrante 5 (Tempo: 2m45s | Acumulado: 29:00)",
        "theme": "couchbase",
        "card1": {
            "title": "🏢 Gigantes em Produção com Couchbase",
            "border": BORDER_GREEN,
            "items": [
                "LinkedIn: Mais de 10 milhões de queries/s com latência < 2ms servindo perfis e conexões.",
                "Amadeus, Sabre & United Airlines: Disponibilidade de voos e cartões de embarque em tempo real.",
                "PayPal & Revolut: Detecção de fraudes em pagamentos com análise de perfil sub-milissegundo.",
                "Domino's Pizza & Tesco: Suporte a picos massivos de pedidos na Black Friday e Super Bowl."
            ]
        },
        "card2": {
            "title": "⚠️ Trade-offs (Quando NÃO Usar)",
            "border": BORDER_ORANGE,
            "items": [
                "Alto consumo de hardware: Exige servidores com dezenas de GBs de RAM dedicados.",
                "Inviável para pequenos microsserviços em containers ou VMs modestas (512 MB a 1 GB).",
                "Complexidade operacional em Day-2 Ops: Dimensionamento de cotas MDS e rebalance de vBuckets.",
                "Custo de licenciamento da versão Enterprise para recursos como Magma e MOI."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 5 (2 minutos e 45 segundos):\n\n"
            "Cases globais demonstram o poder do Couchbase: o LinkedIn atende 10 milhões de operações por segundo, e empresas de aviação como Amadeus processam voos no mundo inteiro com SLAs rigorosos.\n\n"
            "Porém, o trade-off é o custo de infraestrutura: o Couchbase precisa de muita memória RAM para manter o working set aquecido e demanda equipe experiente para gerenciar rebalances. Passo a palavra ao Integrante 6 para a matriz comparativa."
        )
    },
    {
        "meta": "SLIDE 14 • MATRIZ COMPARATIVA DIRETA (PARTE 1/2)",
        "title": "Quadro Comparativo: Modelo de Dados, CAP e Consultas",
        "subtitle": "Confronto direto das garantias distribuídas e linguagens de manipulação",
        "speaker": "Integrante 6 (Tempo: 2m45s | Acumulado: 31:45)",
        "theme": "mixed",
        "card1": {
            "title": "🟧 Apache CouchDB",
            "border": BORDER_ORANGE,
            "items": [
                "Modelo: Documentos JSON + anexos binários em Databases planos sem schemas.",
                "Teorema CAP: AP (PC/EC) — Prioriza disponibilidade imediata e consistência eventual.",
                "Concorrência: MVCC otimista explícito via árvore _rev; sem transações multi-documento.",
                "Consultas: Views MapReduce pré-computadas e Mango Queries (_find). Não suporta JOINs."
            ]
        },
        "card2": {
            "title": "🟦 Couchbase Server",
            "border": BORDER_BLUE,
            "items": [
                "Modelo: JSON + Chave-Valor em hierarquia Bucket ➔ Scope ➔ Collection.",
                "Teorema CAP: CP por padrão (PC/EC) — Único vBucket ativo garante consistência imediata.",
                "Concorrência: CAS otimista + Locks + Transações ACID Distribuídas Multi-Documento.",
                "Consultas: SQL++ (N1QL) com ANSI JOINs, UNNEST, CTEs + Full-Text e Vector Search."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 6 (2 minutos e 45 segundos):\n\n"
            "Obrigado. Consolidamos nossa análise em uma matriz comparativa. Primeiro, no modelo de dados: o CouchDB usa bancos planos, enquanto o Couchbase organiza em Bucket, Scope e Collection.\n\n"
            "No Teorema CAP: o CouchDB é clássico AP, aceitando escritas em qualquer nó desconectado. Já o Couchbase é CP por padrão, pois cada documento é servido por um único vBucket ativo. Nas consultas, o Couchbase se destaca com SQL++ e JOINs nativos."
        )
    },
    {
        "meta": "SLIDE 15 • MATRIZ COMPARATIVA DIRETA (PARTE 2/2)",
        "title": "Quadro Comparativo: Replicação, Latência e Operação",
        "subtitle": "Critérios objetivos de decisão técnica para Arquitetos de Software",
        "speaker": "Integrante 6 (Tempo: 2m45s | Acumulado: 34:30)",
        "theme": "mixed",
        "card1": {
            "title": "🟧 Apache CouchDB",
            "border": BORDER_ORANGE,
            "items": [
                "Replicação: Couch Replication Protocol via HTTP (_changes). Perfeito para P2P e PouchDB.",
                "Performance: 5 a 30 ms por operação, limitada por I/O append-only e overhead HTTP.",
                "Operação: Binário leve em Erlang; requer atenção à compactação de disco e conflitos _rev.",
                "Ideal para: Ambientes remotos, IoT, edge e sincronização móvel offline."
            ]
        },
        "card2": {
            "title": "🟦 Couchbase Server",
            "border": BORDER_BLUE,
            "items": [
                "Replicação: DCP (intra-cluster) e XDCR (inter-cluster) via stream binário TCP de RAM para RAM.",
                "Performance: Sub-milissegundo (< 1 ms / microssegundos) em RAM para operações de chave.",
                "Operação: Plataforma robusta que exige dimensionamento de RAM e nós MDS especializados.",
                "Ideal para: Aplicações de missão crítica com altíssima vazão e baixa latência contínua."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 6 (2 minutos e 45 segundos):\n\n"
            "Nos pilares finais da comparação: a replicação do CouchDB é universal sobre HTTP, integrando navegadores com facilidade. O Couchbase foca em protocolos binários TCP diretos da memória RAM com latência em microssegundos.\n\n"
            "Em resumo: escolha CouchDB quando a conectividade for o desafio principal; escolha Couchbase quando a escala e a latência extrema forem o requisito. Passo agora ao Integrante 7 para o Workshop Prático."
        )
    },
    {
        "meta": "SLIDE 16 • WORKSHOP PRÁTICO: ARQUITETURA DOCKER",
        "title": "Arquitetura do Laboratório Prático (Docker Compose)",
        "subtitle": "Subindo Apache CouchDB 3.3 e Couchbase Server 7.6 no mesmo ambiente de teste",
        "speaker": "Integrante 7 (Tempo: 2 min | Acumulado: 36:30)",
        "theme": "mixed",
        "card1": {
            "title": "🐳 Container Apache CouchDB 3.3",
            "border": BORDER_ORANGE,
            "items": [
                "Imagem oficial: apache/couchdb:3.3 (porta 5984 exposta).",
                "Ambiente: admin / admin123 com inicialização single-node via cURL.",
                "Interface gráfica web Fauxton integrada acessível em /_utils.",
                "Foco dos testes: API HTTP, conflitos MVCC e monitoramento do stream _changes."
            ]
        },
        "card2": {
            "title": "🐳 Container Couchbase Server 7.6",
            "border": BORDER_BLUE,
            "items": [
                "Imagem oficial: couchbase:community-7.6.0 (portas 8091, 8093 e 11210).",
                "Serviços inicializados via CLI: Data, Index e Query.",
                "Console web na porta 8091 com Query Workbench interativo.",
                "Foco dos testes: Criação de Scopes/Collections, índices GSI e queries SQL++."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 7 (2 minutos):\n\n"
            "Obrigado. Vamos provar esses conceitos na prática. Preparamos um laboratório 100% reproduzível via Docker Compose orquestrando o Apache CouchDB e o Couchbase Server.\n\n"
            "Notem a simplicidade de portas: o CouchDB precisa apenas da 5984 para API e interface Fauxton. Já o Couchbase expõe a 8091 para o console web, a 8093 para o motor SQL++ e a 11210 para o acesso binário aos vBuckets."
        )
    },
    {
        "meta": "SLIDE 17 • LIVE DEMO: APACHE COUCHDB",
        "title": "Prática 1: Apache CouchDB via cURL, MVCC e _changes",
        "subtitle": "Demonstração ao vivo do bloqueio otimista HTTP 409 Conflict e streaming de mutações",
        "speaker": "Integrante 7 (Tempo: 2m15s | Acumulado: 38:45)",
        "theme": "couchdb",
        "card1": {
            "title": "💻 Ciclo MVCC & Conflito 409",
            "border": BORDER_ORANGE,
            "items": [
                "PUT /universidade_db/aluno_01 ➔ Criação do registro com revisão 1-4c61...",
                "PUT /universidade_db/aluno_01 (sem _rev) ➔ Dispara erro HTTP 409 Conflict!",
                "PUT enviando _rev atualizado ➔ Sucesso e avanço determinístico para revisão 2-7051...",
                "Validação prática do bloqueio otimista sem locks pessimistas no banco."
            ]
        },
        "card2": {
            "title": "🥭 Mango Query & Stream _changes",
            "border": BORDER_GREEN,
            "items": [
                "POST /_find com seletor {\"cra\": {\"$gt\": 8.5}} para busca declarativa em JSON.",
                "GET /_changes?feed=continuous em terminal secundário monitorando mutações ao vivo.",
                "Inserção de novos documentos salta instantaneamente no stream contínuo.",
                "Base de funcionamento do protocolo de replicação distribuída do CouchDB."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 7 (2 minutos e 15 segundos):\n\n"
            "Na primeira prática, criamos o documento aluno_01 no CouchDB. Quando tentamos atualizar a nota sem passar o _rev atual, o servidor devolve na hora o erro HTTP 409 Document update conflict, provando o MVCC ao vivo.\n\n"
            "Ao passar o _rev correto, a atualização é aceita. Em seguida, mostramos o Mango Query com filtro por CRA e o feed contínuo do endpoint _changes exibindo os deltas em tempo real."
        )
    },
    {
        "meta": "SLIDE 18 • LIVE DEMO: COUCHBASE SQL++ E SCRIPT CRUD",
        "title": "Prática 2: Couchbase N1QL, GSI e Automação Python",
        "subtitle": "Execução de consultas relacionais sobre JSON e fluxo CRUD com tratamento de concorrência",
        "speaker": "Integrante 7 (Tempo: 2m15s | Acumulado: 41:00)",
        "theme": "couchbase",
        "card1": {
            "title": "⚡ Consultas SQL++ no Query Workbench",
            "border": BORDER_BLUE,
            "items": [
                "Criação de índice GSI secundário: idx_curso_cra nos campos curso e cra DESC.",
                "Consulta SQL++ com operador UNNEST para achatar array interno de disciplinas.",
                "Execução em menos de 2 milissegundos com 0% Cache Miss Ratio (100% RAM).",
                "Inspeção visual do plano de execução no Web Console da porta 8091."
            ]
        },
        "card2": {
            "title": "🐍 Automação CRUD em Python 3",
            "border": BORDER_GREEN,
            "items": [
                "Script crud_couchdb_demo.py executando o ciclo completo de persistência.",
                "Tratamento programático do erro HTTP 409 capturando e reenviando a revisão correta.",
                "Deleção segura através da geração de Tombstone (\"_deleted\": true).",
                "Conclusão da apresentação abrindo para a dinâmica de avaliação da turma."
            ]
        },
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 7 (2 minutos e 15 segundos):\n\n"
            "No Couchbase, entramos no Query Workbench, criamos um índice secundário GSI e rodamos um SELECT com UNNEST no array de disciplinas. O resultado volta em 1.4 milissegundos, servido diretamente da memória RAM.\n\n"
            "Por fim, executamos nosso script Python demonstrando a automação de um fluxo profissional com captura e recuperação de conflitos MVCC. Concluímos assim o conteúdo principal!"
        )
    }
]


def add_slide(prs, data):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 1. Background com preenchimento sólido escuro profundo
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR

    # 2. Barra Superior Estilizada (Header HUD)
    top_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.58)
    )
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = TOPBAR_BG
    top_bar.line.color.rgb = BORDER_DEFAULT
    top_bar.line.width = Pt(1)

    tf_top = top_bar.text_frame
    tf_top.word_wrap = True
    p_top = tf_top.paragraphs[0]
    p_top.text = f"  {data['meta']}   |   🎙 {data['speaker']}"
    p_top.font.size = Pt(11.5)
    p_top.font.bold = True
    p_top.font.color.rgb = BORDER_BLUE

    # 3. Faixa de Destaque Decorativa no Topo
    accent_strip = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0.55), Inches(13.333), Inches(0.04)
    )
    accent_strip.fill.solid()
    if data["theme"] == "couchdb":
        accent_strip.fill.fore_color.rgb = BORDER_ORANGE
    elif data["theme"] == "couchbase":
        accent_strip.fill.fore_color.rgb = BORDER_BLUE
    else:
        accent_strip.fill.fore_color.rgb = BORDER_GREEN
    accent_strip.line.fill.background()

    # 4. Título Principal
    title_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.72), Inches(12.5), Inches(0.75))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    tf_title.margin_left = Inches(0)
    p_title = tf_title.paragraphs[0]
    p_title.text = data["title"]
    p_title.font.size = Pt(29)
    p_title.font.bold = True
    p_title.font.color.rgb = TEXT_WHITE

    # 5. Subtítulo Explicativo
    sub_box = slide.shapes.add_textbox(Inches(0.4), Inches(1.42), Inches(12.5), Inches(0.45))
    tf_sub = sub_box.text_frame
    tf_sub.word_wrap = True
    tf_sub.margin_left = Inches(0)
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = data["subtitle"]
    p_sub.font.size = Pt(15.5)
    p_sub.font.color.rgb = TEXT_MUTED

    # 6. Card Esquerdo
    c1_data = data["card1"]
    card1 = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(2.05), Inches(6.1), Inches(5.15)
    )
    card1.fill.solid()
    card1.fill.fore_color.rgb = CARD_BG
    card1.line.color.rgb = c1_data["border"]
    card1.line.width = Pt(2.5)

    tf1 = card1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = Inches(0.3)
    tf1.margin_right = Inches(0.3)
    tf1.margin_top = Inches(0.28)

    p1_head = tf1.paragraphs[0]
    p1_head.text = c1_data["title"]
    p1_head.font.size = Pt(21)
    p1_head.font.bold = True
    p1_head.font.color.rgb = c1_data["border"]
    p1_head.space_after = Pt(14)

    for item in c1_data["items"]:
        p = tf1.add_paragraph()
        p.text = f"▸  {item}"
        p.font.size = Pt(15.5)
        p.font.color.rgb = TEXT_WHITE
        p.space_after = Pt(11)

    # 7. Card Direito
    c2_data = data["card2"]
    card2 = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.83), Inches(2.05), Inches(6.1), Inches(5.15)
    )
    card2.fill.solid()
    card2.fill.fore_color.rgb = CARD_BG
    card2.line.color.rgb = c2_data["border"]
    card2.line.width = Pt(2.5)

    tf2 = card2.text_frame
    tf2.word_wrap = True
    tf2.margin_left = Inches(0.3)
    tf2.margin_right = Inches(0.3)
    tf2.margin_top = Inches(0.28)

    p2_head = tf2.paragraphs[0]
    p2_head.text = c2_data["title"]
    p2_head.font.size = Pt(21)
    p2_head.font.bold = True
    p2_head.font.color.rgb = c2_data["border"]
    p2_head.space_after = Pt(14)

    for item in c2_data["items"]:
        p = tf2.add_paragraph()
        p.text = f"▸  {item}"
        p.font.size = Pt(15.5)
        p.font.color.rgb = TEXT_WHITE
        p.space_after = Pt(11)

    # 8. Roteiro de Fala nas Notas do Apresentador
    notes_slide = slide.notes_slide
    notes_tf = notes_slide.notes_text_frame
    notes_tf.text = data["notes"]


def add_cover_slide(prs):
    """Cria uma capa cinematográfica e moderna no PowerPoint."""
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 1. Background Sólido Escuro Profundo
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR

    # 2. Barra Superior com Gradiente Sutil de Borda
    accent_strip = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.06)
    )
    accent_strip.fill.solid()
    accent_strip.fill.fore_color.rgb = BORDER_BLUE
    accent_strip.line.fill.background()

    # 3. Badge Institucional Centralizado
    badge = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.2), Inches(0.65), Inches(6.933), Inches(0.48)
    )
    badge.fill.solid()
    badge.fill.fore_color.rgb = TOPBAR_BG
    badge.line.color.rgb = BORDER_BLUE
    badge.line.width = Pt(1.5)
    tf_b = badge.text_frame
    tf_b.word_wrap = True
    p_b = tf_b.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    p_b.text = "⚡ SEMINÁRIO DE BANCO DE DADOS AVANÇADO • ENGENHARIA DE SOFTWARE"
    p_b.font.size = Pt(11)
    p_b.font.bold = True
    p_b.font.color.rgb = ACCENT_CYAN

    # 4. Título Principal em Destaque
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.3), Inches(12.133), Inches(1.4))
    tf_t = title_box.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    p_t.alignment = PP_ALIGN.CENTER
    p_t.text = "Apache CouchDB  vs.  Couchbase Server"
    p_t.font.size = Pt(40)
    p_t.font.bold = True
    p_t.font.color.rgb = TEXT_WHITE

    # 5. Tagline Explicativa
    tag_box = slide.shapes.add_textbox(Inches(1.2), Inches(2.35), Inches(10.933), Inches(0.6))
    tf_tag = tag_box.text_frame
    tf_tag.word_wrap = True
    p_tag = tf_tag.paragraphs[0]
    p_tag.alignment = PP_ALIGN.CENTER
    p_tag.text = "Da Sincronização Offline-First sobre HTTP à Performance Sub-Milissegundo em Memória Gerenciada"
    p_tag.font.size = Pt(16.5)
    p_tag.font.color.rgb = TEXT_MUTED

    # 6. Duelo Visual: Card CouchDB (Esquerda)
    c1 = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(3.1), Inches(5.35), Inches(2.7)
    )
    c1.fill.solid()
    c1.fill.fore_color.rgb = CARD_BG
    c1.line.color.rgb = BORDER_ORANGE
    c1.line.width = Pt(2.5)
    tf_c1 = c1.text_frame
    tf_c1.word_wrap = True
    tf_c1.margin_left = Inches(0.25)
    tf_c1.margin_top = Inches(0.2)

    p_h1 = tf_c1.paragraphs[0]
    p_h1.text = "🟧 Apache CouchDB (AP • REST • MVCC)"
    p_h1.font.size = Pt(17)
    p_h1.font.bold = True
    p_h1.font.color.rgb = BORDER_ORANGE
    p_h1.space_after = Pt(8)

    for item in [
        "Motor Append-Only B+Tree imune a corrupções.",
        "API 100% HTTP/REST nativa (porta 5984).",
        "Sincronização Offline-First na borda com PouchDB."
    ]:
        p = tf_c1.add_paragraph()
        p.text = f"▸  {item}"
        p.font.size = Pt(13.5)
        p.font.color.rgb = TEXT_WHITE
        p.space_after = Pt(6)

    # 7. Duelo Visual: Card Couchbase (Direita)
    c2 = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.983), Inches(3.1), Inches(5.35), Inches(2.7)
    )
    c2.fill.solid()
    c2.fill.fore_color.rgb = CARD_BG
    c2.line.color.rgb = BORDER_BLUE
    c2.line.width = Pt(2.5)
    tf_c2 = c2.text_frame
    tf_c2.word_wrap = True
    tf_c2.margin_left = Inches(0.25)
    tf_c2.margin_top = Inches(0.2)

    p_h2 = tf_c2.paragraphs[0]
    p_h2.text = "🟦 Couchbase Server (CP • RAM • SQL++)"
    p_h2.font.size = Pt(17)
    p_h2.font.bold = True
    p_h2.font.color.rgb = BORDER_BLUE
    p_h2.space_after = Pt(8)

    for item in [
        "Arquitetura Memory-First com latência sub-ms.",
        "Particionamento por 1024 vBuckets (Single-Hop).",
        "Linguagem N1QL (SQL++) com ANSI JOINs e GSI."
    ]:
        p = tf_c2.add_paragraph()
        p.text = f"▸  {item}"
        p.font.size = Pt(13.5)
        p.font.color.rgb = TEXT_WHITE
        p.space_after = Pt(6)

    # 8. Faixa da Equipe (Rodapé da Capa)
    team_bar = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(6.1), Inches(11.333), Inches(0.85)
    )
    team_bar.fill.solid()
    team_bar.fill.fore_color.rgb = TOPBAR_BG
    team_bar.line.color.rgb = BORDER_DEFAULT
    team_bar.line.width = Pt(1)
    tf_team = team_bar.text_frame
    tf_team.word_wrap = True
    tf_team.margin_top = Inches(0.12)
    p_team_title = tf_team.paragraphs[0]
    p_team_title.alignment = PP_ALIGN.CENTER
    p_team_title.text = "EQUIPE DE APRESENTAÇÃO (7 INTEGRANTES — 35 A 45 MINUTOS)"
    p_team_title.font.size = Pt(10.5)
    p_team_title.font.bold = True
    p_team_title.font.color.rgb = ACCENT_YELLOW
    p_team_title.space_after = Pt(4)

    p_members = tf_team.add_paragraph()
    p_members.alignment = PP_ALIGN.CENTER
    p_members.text = "Int. 1: Origem & Filosofia  |  Int. 2: Motor CouchDB  |  Int. 3: Motor Couchbase  |  Int. 4: Mercado CouchDB  |  Int. 5: Mercado Couchbase  |  Int. 6: Matriz  |  Int. 7: Workshop"
    p_members.font.size = Pt(10)
    p_members.font.color.rgb = TEXT_MUTED

    # Notas do Apresentador para a Capa
    notes_slide = slide.notes_slide
    notes_tf = notes_slide.notes_text_frame
    notes_tf.text = (
        "CAPA OFICIAL DO SEMINÁRIO:\n"
        "Apresentação da equipe de 7 integrantes para a disciplina de Bancos de Dados / Sistemas Distribuídos.\n"
        "O Integrante 1 inicia a apresentação com a introdução do problema da Web 2.0 e a crise dos bancos relacionais tradicionais."
    )


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # 1. Slide de Capa Oficial
    add_cover_slide(prs)

    # 2. Slides do Conteúdo
    for slide_data in SLIDES_DATA:
        add_slide(prs, slide_data)

    output_path = r"d:\Área de Trabalho\BDChrouch\Apresentacao_CouchDB_vs_Couchbase.pptx"
    prs.save(output_path)
    print(f"PPTX gerado com sucesso em: {output_path}")


if __name__ == "__main__":
    main()
